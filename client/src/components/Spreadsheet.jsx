import React, { useRef, useEffect, useCallback, useMemo, useState } from 'react';
import { HotTable } from '@handsontable/react-wrapper';
import Handsontable from 'handsontable/base';
import { registerAllModules } from 'handsontable/registry';
import { getSheetWindow, updateCell } from '../services/api';

// Register all Handsontable modules
registerAllModules();

/**
 * Spreadsheet - Handsontable-based Excel viewer/editor with virtualized lazy loading
 * 
 * Features:
 * - Excel-style smooth scrolling
 * - Virtualized rendering with window replacement (NOT appending)
 * - Multi-sheet support
 * - Lazy loading on scroll (vertical and horizontal)
 * - Edit synchronization with backend
 * - Constant memory usage
 */
const Spreadsheet = ({ filePath, sheetId, onDataChange }) => {
  const hotRef = useRef(null);
  const workbookIdRef = useRef(null);
  
  // Data state - stores ONLY the current window
  const [data, setData] = useState([[]]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  
  // Metadata and window tracking
  const metadataRef = useRef(null);
  const currentWindowRef = useRef({ rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 });
  const loadingRef = useRef(false);
  const scrollDebounceTimerRef = useRef(null);
  const pendingEditTimerRef = useRef(null);
  const pendingEditsRef = useRef([]);
  
  // Constants
  const SCROLL_BUFFER = 30; // Rows/cols to load beyond visible area
  const SCROLL_DEBOUNCE_MS = 150; // Debounce scroll requests
  const EDIT_DEBOUNCE_MS = 1000; // Debounce edit sync

  /**
   * Extract workbookId from filePath
   * Note: Don't encode here - api.js will handle encoding
   */
  const workbookId = useMemo(() => {
    if (!filePath) return null;
    return filePath; // Pass raw filePath, api.js will encode it
  }, [filePath]);

  /**
   * Fetch and REPLACE data window (NOT append)
   * This is the core of virtualized loading - we replace the entire data array
   * with only the visible window + buffer
   */
  const fetchDataWindow = useCallback(async (rowStart, rowEnd, colStart, colEnd, force = false) => {
    if (!workbookId || loadingRef.current) return;
    
    // Check if we're requesting the same window (prevent duplicate requests)
    const currentWindow = currentWindowRef.current;
    if (!force && 
        currentWindow.rowStart === rowStart && 
        currentWindow.rowEnd === rowEnd &&
        currentWindow.colStart === colStart &&
        currentWindow.colEnd === colEnd) {
      return;
    }

    try {
      loadingRef.current = true;
      setLoading(true);
      setError(null);
      
      console.log('🔄 Fetching window:', { rowStart, rowEnd, colStart, colEnd, sheetId });
      
      if (!sheetId) {
        throw new Error('Sheet ID (name) is required');
      }
      
      const response = await getSheetWindow(workbookId, rowStart, rowEnd, colStart, colEnd, sheetId);
      const { data: responseData, meta } = response.data;

      // Store metadata
      if (meta) {
        metadataRef.current = meta;
      }

      // CRITICAL: Replace entire data array with window data
      // This ensures we never accumulate data - memory stays constant
      if (responseData && Array.isArray(responseData)) {
        // Create a full-size sparse array for the entire sheet
        // But only populate the window we fetched
        const totalRows = meta?.totalRows || Math.max(rowEnd, 1000);
        const totalCols = meta?.totalColumns || Math.max(colEnd, 26);
        
        // Initialize full array with nulls
        const newData = Array(totalRows).fill(null).map(() => Array(totalCols).fill(null));
        
        // Fill only the window we fetched
        responseData.forEach((row, rowIdx) => {
          const actualRow = rowStart - 1 + rowIdx; // Convert to 0-based
          if (actualRow >= 0 && actualRow < newData.length && Array.isArray(row)) {
            row.forEach((cell, colIdx) => {
              const actualCol = colStart - 1 + colIdx; // Convert to 0-based
              if (actualCol >= 0 && actualCol < newData[actualRow].length) {
                newData[actualRow][actualCol] = cell;
              }
            });
          }
        });

        // REPLACE data (not merge/append)
        setData(newData);
        currentWindowRef.current = { rowStart, rowEnd, colStart, colEnd };
        
        console.log('✅ Window loaded:', { 
          rows: newData.length, 
          cols: newData[0]?.length || 0,
          window: { rowStart, rowEnd, colStart, colEnd }
        });
      }
      
      setLoading(false);
    } catch (error) {
      console.error('❌ Error fetching window:', error);
      setError(error.response?.data?.message || error.message || 'Failed to load spreadsheet data');
      setLoading(false);
      setData([[]]);
    } finally {
      loadingRef.current = false;
    }
  }, [workbookId, sheetId]);

  // Track previous viewport to detect scroll direction
  const previousViewportRef = useRef({ row: 0, col: 0 });

  /**
   * Handle scroll - lazy load both rows and columns
   * Uses afterScroll hook which fires on any scroll
   */
  const handleScroll = useCallback(() => {
    if (!hotRef.current?.hotInstance || !metadataRef.current || loadingRef.current) return;

    const instance = hotRef.current.hotInstance;
    
    try {
      // Get current viewport
      const viewport = instance.view.getViewport();
      const firstVisibleRow = viewport[0];
      const lastVisibleRow = viewport[2];
      const firstVisibleCol = viewport[1];
      const lastVisibleCol = viewport[3];
      
      if (firstVisibleRow === null || lastVisibleRow === null || 
          firstVisibleCol === null || lastVisibleCol === null) {
        return;
      }
      
      // Check if viewport actually changed
      const prevViewport = previousViewportRef.current;
      if (prevViewport.row === firstVisibleRow && prevViewport.col === firstVisibleCol) {
        return; // No change, skip
      }
      
      previousViewportRef.current = { row: firstVisibleRow, col: firstVisibleCol };
      
      // Calculate window with buffer for both dimensions
      const rowStart = Math.max(1, firstVisibleRow - SCROLL_BUFFER + 1); // Convert to 1-based
      const rowEnd = Math.min(
        metadataRef.current.totalRows, 
        lastVisibleRow + SCROLL_BUFFER + 1
      );
      
      const colStart = Math.max(1, firstVisibleCol - SCROLL_BUFFER + 1); // Convert to 1-based
      const colEnd = Math.min(
        metadataRef.current.totalColumns,
        lastVisibleCol + SCROLL_BUFFER + 1
      );
      
      // Debounce scroll requests
      if (scrollDebounceTimerRef.current) {
        clearTimeout(scrollDebounceTimerRef.current);
      }
      
      scrollDebounceTimerRef.current = setTimeout(() => {
        fetchDataWindow(rowStart, rowEnd, colStart, colEnd);
      }, SCROLL_DEBOUNCE_MS);
      
    } catch (error) {
      console.error('Error in scroll handler:', error);
    }
  }, [fetchDataWindow]);

  /**
   * Handle cell changes - sync to backend
   */
  const handleAfterChange = useCallback((changes, source) => {
    // CRITICAL: Ignore programmatic changes to prevent loops
    if (!changes || source === 'loadData' || source === 'updateData' || source === 'CopyPaste.paste') {
      return;
    }

    if (!workbookId || !hotRef.current?.hotInstance) return;

    const instance = hotRef.current.hotInstance;
    
    // Collect changes
    changes.forEach(([row, col, oldValue, newValue]) => {
      if (oldValue !== newValue) {
        pendingEditsRef.current.push({
          row: row + 1, // Convert to 1-based for backend
          col: col + 1,
          value: newValue
        });
      }
    });

    // Debounce edit sync
    if (pendingEditTimerRef.current) {
      clearTimeout(pendingEditTimerRef.current);
    }

    pendingEditTimerRef.current = setTimeout(async () => {
      const edits = [...pendingEditsRef.current];
      pendingEditsRef.current = [];

      // Sync each edit to backend
      for (const edit of edits) {
        try {
          if (!sheetId) {
            throw new Error('Sheet ID (name) is required for cell updates');
          }
          await updateCell(workbookId, sheetId, edit.row, edit.col, edit.value);
          console.log('✅ Cell synced:', edit);
        } catch (error) {
          console.error('❌ Failed to sync cell:', edit, error);
          // Re-add failed edit to retry queue
          pendingEditsRef.current.push(edit);
        }
      }

      // Notify parent component
      if (onDataChange && edits.length > 0) {
        onDataChange(edits);
      }
    }, EDIT_DEBOUNCE_MS);
  }, [workbookId, sheetId, onDataChange]);

  /**
   * Initial load - fetch first window
   */
  useEffect(() => {
    if (!workbookId) {
      setData([[]]);
      setLoading(false);
      return;
    }

    workbookIdRef.current = workbookId;
    
    // Reset state when workbook or sheet changes
    setData([[]]);
    setLoading(true);
    setError(null);
    metadataRef.current = null;
    currentWindowRef.current = { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 };
    loadingRef.current = false;

    // Load initial window
    const loadInitial = async () => {
      if (sheetId) {
        await fetchDataWindow(1, 100, 1, 30, true);
      }
    };

    loadInitial();
  }, [workbookId, sheetId, fetchDataWindow]);

  /**
   * Update Handsontable when data changes
   */
  useEffect(() => {
    if (!hotRef.current?.hotInstance || data.length === 0) return;

    const instance = hotRef.current.hotInstance;
    
    // Only update if data actually changed
    const currentData = instance.getData();
    if (JSON.stringify(currentData) !== JSON.stringify(data)) {
      // Use loadData to replace entire dataset
      instance.loadData(data);
      console.log('📊 Data loaded into Handsontable');
    }
  }, [data]);

  // Calculate container height
  const [containerHeight, setContainerHeight] = useState(() => 
    typeof window !== 'undefined' ? window.innerHeight - 120 : 600
  );

  useEffect(() => {
    const handleResize = () => {
      setContainerHeight(window.innerHeight - 120);
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  // Cleanup
  useEffect(() => {
    return () => {
      if (scrollDebounceTimerRef.current) {
        clearTimeout(scrollDebounceTimerRef.current);
      }
      if (pendingEditTimerRef.current) {
        clearTimeout(pendingEditTimerRef.current);
      }
    };
  }, []);

  if (!workbookId) {
    return (
      <div className="flex items-center justify-center h-full text-slate-500">
        No file selected
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex flex-col items-center justify-center h-full text-red-500">
        <p className="mb-2">Error loading spreadsheet:</p>
        <p className="text-sm text-slate-400">{error}</p>
      </div>
    );
  }

  const displayData = data.length > 0 && data[0]?.length > 0 ? data : [[]];

  return (
    <div
      id="sheet-container"
      className="w-full"
      style={{
        height: `${containerHeight}px`,
        width: '100%',
        position: 'relative',
        overflow: 'hidden',
        minHeight: '400px'
      }}
    >
      {loading && (!data.length || (data.length === 1 && !data[0]?.length)) && (
        <div className="absolute inset-0 flex items-center justify-center bg-white/50 dark:bg-slate-900/50 z-10">
          <div className="text-slate-500">Loading spreadsheet data...</div>
        </div>
      )}
      <HotTable
        key={`${workbookId}-${sheetId}`} // Force re-render when workbook or sheet changes
        ref={hotRef}
        data={displayData}
        colHeaders={true}
        rowHeaders={true}
        width="100%"
        height={containerHeight}
        rowHeights={28}
        colWidths={100}
        autoRowSize={false}
        autoColumnSize={false}
        manualRowResize={true}
        manualColumnResize={true}
        renderAllRows={false}
        viewportRowRenderingOffset={SCROLL_BUFFER}
        viewportColumnRenderingOffset={SCROLL_BUFFER}
        licenseKey="non-commercial-and-evaluation"
        themeName="ht-theme-main"
        afterChange={handleAfterChange}
        afterScrollHorizontally={handleScroll}
        afterScrollVertically={handleScroll}
        afterInit={() => {
          console.log('✅ Handsontable initialized');
          if (hotRef.current?.hotInstance) {
            const instance = hotRef.current.hotInstance;
            console.log('Instance info:', {
              rows: instance.countRows(),
              cols: instance.countCols()
            });
          }
        }}
      />
    </div>
  );
};

export default React.memo(Spreadsheet);
