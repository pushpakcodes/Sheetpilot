import React, { useRef, useEffect, useCallback, useMemo, useState } from 'react';
import { HotTable } from '@handsontable/react-wrapper';
import Handsontable from 'handsontable/base';
import { registerAllModules } from 'handsontable/registry';
import { getSheetWindow } from '../services/api';

// Register all Handsontable modules
registerAllModules();

/**
 * Spreadsheet - Handsontable-based Excel viewer/editor
 * 
 * Features:
 * - Excel-like smooth scrolling
 * - Virtualized rendering for large datasets
 * - Editable cells with backend sync
 * - Proper grid alignment
 * - Fixed-size container for stable rendering
 */
const Spreadsheet = ({ filePath, sheetIndex = 0, onDataChange }) => {
  const hotRef = useRef(null);
  const [data, setData] = useState([[]]); // Initialize with at least one empty row
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const metadataRef = useRef(null);
  const loadingRef = useRef(false);
  const pendingChangesRef = useRef([]);
  const debounceTimerRef = useRef(null);
  const isMountedRef = useRef(false);

  /**
   * Extract filename from filePath and URL encode it
   */
  const sheetId = useMemo(() => {
    if (!filePath) return null;
    // URL encode the filePath to handle special characters
    // The backend can handle both encoded paths and filenames
    return encodeURIComponent(filePath);
  }, [filePath]);

  /**
   * Fetch data window from backend
   */
  const fetchDataWindow = useCallback(async (rowStart, rowEnd, colStart, colEnd) => {
    if (!sheetId || loadingRef.current) return;

    try {
      loadingRef.current = true;
      setLoading(true);
      setError(null);
      console.log('Fetching data window:', { sheetId, rowStart, rowEnd, colStart, colEnd, sheetIndex });
      const response = await getSheetWindow(sheetId, rowStart, rowEnd, colStart, colEnd, sheetIndex);
      console.log('API Response:', response.data);
      
      const { data: responseData, meta } = response.data;

      // Store metadata
      if (meta) {
        metadataRef.current = meta;
        console.log('Metadata stored:', meta);
      }

      // Convert 2D array to Handsontable format
      // Backend returns data[row][col], we need to ensure it's properly formatted
      if (responseData && Array.isArray(responseData)) {
        setData(prevData => {
          // Initialize data array if needed
          let newData = prevData;
          if ((!newData.length || newData.length === 1 && !newData[0].length) && meta) {
            const totalRows = Math.max(meta.totalRows || 100, 1);
            const totalCols = Math.max(meta.totalColumns || 26, 1);
            newData = Array(totalRows).fill(null).map(() => Array(totalCols).fill(null));
            console.log('Initialized data array:', { totalRows, totalCols });
          } else if (!newData.length || (newData.length === 1 && !newData[0].length)) {
            // Fallback if no metadata
            newData = Array(100).fill(null).map(() => Array(26).fill(null));
            console.log('Using fallback data array size');
          }

          // Create a copy to avoid mutating state directly
          const updatedData = newData.map(row => [...row]);

          // Update the specific window in our data array
          // Backend data is 0-indexed relative to the window
          responseData.forEach((row, rowIdx) => {
            const actualRow = rowStart - 1 + rowIdx; // Convert to 0-based
            if (actualRow >= 0 && actualRow < updatedData.length) {
              if (Array.isArray(row)) {
                row.forEach((cell, colIdx) => {
                  const actualCol = colStart - 1 + colIdx; // Convert to 0-based
                  if (actualCol >= 0 && actualCol < updatedData[actualRow].length) {
                    updatedData[actualRow][actualCol] = cell;
                  }
                });
              }
            }
          });

          console.log('Data updated, sample:', updatedData.slice(0, 5).map(r => r.slice(0, 5)));
          console.log('Full data dimensions:', { rows: updatedData.length, cols: updatedData[0]?.length || 0 });
          setLoading(false);
          return updatedData;
        });
      } else {
        console.warn('No data array in response:', response.data);
        setLoading(false);
        // Initialize with empty data if no data returned
        if (meta) {
          const totalRows = Math.max(meta.totalRows || 1, 1);
          const totalCols = Math.max(meta.totalColumns || 1, 1);
          setData(Array(totalRows).fill(null).map(() => Array(totalCols).fill(null)));
        }
      }
    } catch (error) {
      console.error('Error fetching data window:', error);
      console.error('Error details:', {
        message: error.message,
        response: error.response?.data,
        status: error.response?.status,
        sheetId
      });
      setError(error.response?.data?.message || error.message || 'Failed to load spreadsheet data');
      setLoading(false);
      // Set to empty array if error occurs
      setData([[]]);
    } finally {
      loadingRef.current = false;
    }
  }, [sheetId, sheetIndex]);

  /**
   * Initial data load - fetch first window to get metadata and initial data
   */
  useEffect(() => {
    if (!sheetId) return;

    // Reset state when sheetId changes
    setData([[]]);
    setLoading(true);
    setError(null);
    metadataRef.current = null;
    loadingRef.current = false;

    const loadInitialData = async () => {
      // Fetch a small window first to get metadata
      await fetchDataWindow(1, 100, 1, 30);
    };

    loadInitialData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheetId]); // fetchDataWindow is stable, so we can safely omit it

  /**
   * Log when Handsontable instance is ready
   */
  useEffect(() => {
    const checkInstance = () => {
      if (hotRef.current?.hotInstance) {
        const instance = hotRef.current.hotInstance;
        console.log('Handsontable instance ready:', {
          rows: instance.countRows(),
          cols: instance.countCols(),
          dataRows: data.length,
          dataCols: data[0]?.length || 0
        });
      }
    };
    
    const timer = setTimeout(checkInstance, 500);
    return () => clearTimeout(timer);
  }, [data.length]);

  /**
   * Mark component as mounted
   */
  useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
    };
  }, []);

  /**
   * Handle cell changes with debouncing
   */
  const handleAfterChange = useCallback((changes, source) => {
    if (!changes || source === 'loadData') return;

    // Store changes for batch update
    changes.forEach(([row, col, oldValue, newValue]) => {
      if (oldValue !== newValue) {
        pendingChangesRef.current.push({
          row: row + 1, // Convert to 1-based for backend
          col: col + 1, // Convert to 1-based for backend
          value: newValue
        });
      }
    });

    // Debounce backend updates
    if (debounceTimerRef.current) {
      clearTimeout(debounceTimerRef.current);
    }

    debounceTimerRef.current = setTimeout(() => {
      if (pendingChangesRef.current.length > 0 && onDataChange) {
        onDataChange(pendingChangesRef.current);
        pendingChangesRef.current = [];
      }
    }, 1000); // 1 second debounce
  }, [onDataChange]);

  /**
   * Handle viewport scrolling - fetch new data windows as needed
   * Using afterScroll callback which fires after scroll events
   */
  const handleAfterScroll = useCallback(() => {
    if (!hotRef.current?.hotInstance || !metadataRef.current || loadingRef.current) return;

    const instance = hotRef.current.hotInstance;
    
    try {
      // Get viewport information using Handsontable API
      const viewport = instance.view.getViewport();
      const firstVisibleRow = viewport[0];
      const lastVisibleRow = viewport[2];
      const firstVisibleCol = viewport[1];
      const lastVisibleCol = viewport[3];
      
      // Fetch data for visible area + buffer
      const rowStart = Math.max(1, firstVisibleRow + 1); // Convert to 1-based
      const rowEnd = Math.min(metadataRef.current.totalRows, lastVisibleRow + 50); // Add buffer
      const colStart = Math.max(1, firstVisibleCol + 1); // Convert to 1-based
      const colEnd = Math.min(metadataRef.current.totalColumns, lastVisibleCol + 20); // Add buffer

      fetchDataWindow(rowStart, rowEnd, colStart, colEnd);
    } catch (error) {
      // Fallback: fetch a reasonable window if API fails
      if (metadataRef.current) {
        fetchDataWindow(1, Math.min(100, metadataRef.current.totalRows), 1, Math.min(30, metadataRef.current.totalColumns));
      }
    }
  }, [fetchDataWindow]);

  if (!sheetId) {
    return (
      <div className="flex items-center justify-center h-full text-slate-500">
        No file selected
      </div>
    );
  }

  // Show loading or error state
  if (error) {
    return (
      <div className="flex flex-col items-center justify-center h-full text-red-500">
        <p className="mb-2">Error loading spreadsheet:</p>
        <p className="text-sm text-slate-400">{error}</p>
      </div>
    );
  }

  // Ensure we have valid data structure
  const displayData = data.length > 0 && data[0]?.length > 0 ? data : [[]];

  // Calculate container height with state to trigger re-renders on resize
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
        key={sheetId} // Force re-render when file changes
        ref={hotRef}
        data={displayData}
        colHeaders={true}
        rowHeaders={true}
        width="100%"
        height={containerHeight}
        stretchH="all"
        rowHeights={28}
        colWidths={120}
        autoRowSize={false}
        autoColumnSize={false}
        manualRowResize={true}
        manualColumnResize={true}
        renderAllRows={false}
        viewportRowRenderingOffset={20}
        viewportColumnRenderingOffset={10}
        licenseKey="non-commercial-and-evaluation"
        themeName="ht-theme-main"
        afterChange={handleAfterChange}
        afterScroll={handleAfterScroll}
        afterInit={() => {
          console.log('Handsontable initialized');
          if (hotRef.current?.hotInstance) {
            const instance = hotRef.current.hotInstance;
            console.log('Instance after init:', {
              rows: instance.countRows(),
              cols: instance.countCols(),
              containerHeight: containerHeight
            });
            // Force render
            instance.render();
          }
        }}
      />
    </div>
  );
};

export default React.memo(Spreadsheet);

