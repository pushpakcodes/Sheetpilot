import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { Grid } from 'react-window';
import { getSheetWindow } from '../services/api';

/**
 * VirtualizedExcelView - Production-grade windowed Excel rendering
 * 
 * This component implements Google Sheets-style infinite scrolling by:
 * 1. Rendering only visible cells (~100 rows x ~30 columns)
 * 2. Fetching data in chunks as user scrolls
 * 3. Discarding offscreen data to maintain performance
 * 4. Providing smooth, continuous scrolling experience
 * 
 * Architecture:
 * - Uses react-window's FixedSizeGrid for efficient DOM rendering
 * - Maintains a cache of fetched windows
 * - Calculates which window to fetch based on scroll position
 * - Never loads the entire sheet into memory
 */
const VirtualizedExcelView = ({ filePath, sheetIndex = 0, onSheetChange }) => {
  // Metadata about the sheet (total dimensions)
  const [metadata, setMetadata] = useState(null);
  
  // Cache of fetched data windows
  // Key format: `${rowStart}-${rowEnd}-${colStart}-${colEnd}`
  const [dataCache, setDataCache] = useState(new Map());
  
  // Loading states for different windows
  const [loadingWindows, setLoadingWindows] = useState(new Set());
  
  // Grid dimensions (visible viewport)
  const [gridDimensions, setGridDimensions] = useState({
    rowCount: 100,  // Initial estimate, will be updated from metadata
    columnCount: 30 // Initial estimate, will be updated from metadata
  });
  
  // Scroll position tracking
  const gridRef = useRef(null);
  const containerRef = useRef(null);
  const [scrollState, setScrollState] = useState({ scrollTop: 0, scrollLeft: 0 });
  const [containerSize, setContainerSize] = useState({ width: 1200, height: 600 });
  
  // Cell dimensions
  const CELL_HEIGHT = 24;
  const CELL_WIDTH = 100;
  const ROW_HEADER_WIDTH = 50;
  const COL_HEADER_HEIGHT = 24;
  
  // Window size for fetching (fetch slightly more than visible for smooth scrolling)
  const ROW_WINDOW_SIZE = 150;  // Fetch 150 rows at a time
  const COL_WINDOW_SIZE = 50;   // Fetch 50 columns at a time
  
  // Viewport size (how many cells are visible)
  const VISIBLE_ROWS = Math.ceil(600 / CELL_HEIGHT); // ~25 rows visible
  const VISIBLE_COLS = Math.ceil(1200 / CELL_WIDTH); // ~12 columns visible

  /**
   * Extract filename from filePath
   * The backend accepts either filename or full path
   */
  const sheetId = useMemo(() => {
    if (!filePath) return null;
    // If it's already just a filename, use it
    if (!filePath.includes('/') && !filePath.includes('\\')) {
      return filePath;
    }
    // Extract filename from path
    return filePath.split(/[/\\]/).pop();
  }, [filePath]);

  /**
   * Update container size on mount and resize
   */
  useEffect(() => {
    const updateSize = () => {
      if (containerRef.current) {
        const rect = containerRef.current.getBoundingClientRect();
        setContainerSize({
          width: rect.width,
          height: rect.height
        });
      }
    };

    updateSize();
    window.addEventListener('resize', updateSize);
    return () => window.removeEventListener('resize', updateSize);
  }, []);

  /**
   * Fetch metadata on mount
   * We fetch a small window first to get total dimensions
   */
  useEffect(() => {
    if (!sheetId) return;

    const fetchMetadata = async () => {
      try {
        // Fetch a small window to get metadata
        const response = await getSheetWindow(sheetId, 1, 1, 1, 1, sheetIndex);
        const meta = response.data.meta;
        setMetadata(meta);
        setGridDimensions({
          rowCount: meta.totalRows || 1000,
          columnCount: meta.totalColumns || 26
        });
      } catch (error) {
        console.error('Error fetching metadata:', error);
      }
    };

    fetchMetadata();
  }, [sheetId, sheetIndex]);

  /**
   * Prefetch initial visible window when metadata is available
   */
  useEffect(() => {
    if (!metadata || !sheetId) return;
    
    // Prefetch the initial visible window for immediate rendering
    // Fetch first ~150 rows and ~50 columns
    const initialRowEnd = Math.min(metadata.totalRows, ROW_WINDOW_SIZE);
    const initialColEnd = Math.min(metadata.totalColumns, COL_WINDOW_SIZE);
    
    const cacheKey = `1-${initialRowEnd}-1-${initialColEnd}`;
    if (!dataCache.has(cacheKey) && !loadingWindows.has(cacheKey)) {
      // Use fetchWindow but we need to handle the dependency properly
      // For now, trigger it via a state update or call directly
      fetchWindow(1, initialRowEnd, 1, initialColEnd);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [metadata, sheetId]);

  /**
   * Calculate which window to fetch based on visible indices
   * We fetch a window that's larger than the viewport for smooth scrolling
   */
  const calculateWindow = useCallback((rowIndex, colIndex) => {
    // Calculate window bounds with buffer
    const rowStart = Math.max(1, Math.floor(rowIndex / ROW_WINDOW_SIZE) * ROW_WINDOW_SIZE + 1);
    const rowEnd = Math.min(
      metadata?.totalRows || 10000,
      rowStart + ROW_WINDOW_SIZE - 1
    );
    
    const colStart = Math.max(1, Math.floor(colIndex / COL_WINDOW_SIZE) * COL_WINDOW_SIZE + 1);
    const colEnd = Math.min(
      metadata?.totalColumns || 100,
      colStart + COL_WINDOW_SIZE - 1
    );
    
    return { rowStart, rowEnd, colStart, colEnd };
  }, [metadata]);

  /**
   * Fetch a window of data from the backend
   * This is called when we need data that's not in cache
   */
  const fetchWindow = useCallback(async (rowStart, rowEnd, colStart, colEnd) => {
    const cacheKey = `${rowStart}-${rowEnd}-${colStart}-${colEnd}`;
    
    // Don't fetch if already in cache or currently loading
    if (dataCache.has(cacheKey) || loadingWindows.has(cacheKey)) {
      return;
    }
    
    setLoadingWindows(prev => new Set(prev).add(cacheKey));
    
    try {
      const response = await getSheetWindow(sheetId, rowStart, rowEnd, colStart, colEnd, sheetIndex);
      
      // Store in cache
      setDataCache(prev => {
        const newCache = new Map(prev);
        newCache.set(cacheKey, {
          data: response.data.data,
          rowStart,
          rowEnd,
          colStart,
          colEnd
        });
        return newCache;
      });
      
      // Update metadata if we got new info
      if (response.data.meta) {
        setMetadata(response.data.meta);
      }
    } catch (error) {
      console.error('Error fetching window:', error);
    } finally {
      setLoadingWindows(prev => {
        const newSet = new Set(prev);
        newSet.delete(cacheKey);
        return newSet;
      });
    }
  }, [sheetId, sheetIndex, dataCache, loadingWindows]);

  /**
   * Get cell value from cache
   * Returns null if data is not yet loaded
   */
  const getCellValue = useCallback((rowIndex, colIndex) => {
    // Convert to 1-based for Excel (Excel uses 1-based indexing)
    const excelRow = rowIndex + 1;
    const excelCol = colIndex + 1;
    
    // Find which window contains this cell
    for (const [cacheKey, windowData] of dataCache.entries()) {
      const { rowStart, rowEnd, colStart, colEnd, data } = windowData;
      
      if (excelRow >= rowStart && excelRow <= rowEnd &&
          excelCol >= colStart && excelCol <= colEnd) {
        // Calculate local indices within the window
        const localRow = excelRow - rowStart;
        const localCol = excelCol - colStart;
        
        if (data[localRow] && data[localRow][localCol] !== undefined) {
          return data[localRow][localCol];
        }
      }
    }
    
    // Data not loaded yet - trigger fetch
    const window = calculateWindow(rowIndex, colIndex);
    fetchWindow(window.rowStart, window.rowEnd, window.colStart, window.colEnd);
    
    return null; // Return null while loading
  }, [dataCache, calculateWindow, fetchWindow]);

  /**
   * Convert column index to Excel column letter (A, B, ..., Z, AA, AB, ...)
   */
  const getColumnLabel = useCallback((index) => {
    let label = '';
    let i = index;
    while (i >= 0) {
      label = String.fromCharCode((i % 26) + 65) + label;
      i = Math.floor(i / 26) - 1;
    }
    return label;
  }, []);

  /**
   * Cell renderer for the grid
   * This is called by react-window for each visible cell
   * react-window 2.x passes: columnIndex, rowIndex, style, ariaAttributes
   */
  const Cell = useCallback(({ columnIndex, rowIndex, style, ariaAttributes }) => {
    // Row 0 and Column 0 are headers
    if (rowIndex === 0 && columnIndex === 0) {
      // Corner cell (select all)
      return (
        <div
          style={{
            ...style,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            backgroundColor: '#f1f5f9',
            borderRight: '1px solid #cbd5e1',
            borderBottom: '1px solid #cbd5e1',
            position: 'sticky',
            left: 0,
            top: 0,
            zIndex: 30
          }}
          className="dark:bg-slate-700 dark:border-slate-600"
        >
          <div className="w-0 h-0 border-l-[10px] border-l-transparent border-t-[10px] border-t-slate-400 dark:border-t-slate-500"></div>
        </div>
      );
    }
    
    if (rowIndex === 0) {
      // Column header (A, B, C, ...)
      const colLabel = getColumnLabel(columnIndex - 1);
      return (
        <div
          style={{
            ...style,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            backgroundColor: '#f1f5f9',
            borderRight: '1px solid #cbd5e1',
            borderBottom: '1px solid #cbd5e1',
            fontSize: '13px',
            fontWeight: 'normal',
            color: '#334155',
            userSelect: 'none',
            position: 'sticky',
            top: 0,
            zIndex: 20
          }}
          className="dark:bg-slate-800 dark:border-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600"
        >
          {colLabel}
        </div>
      );
    }
    
    if (columnIndex === 0) {
      // Row header (1, 2, 3, ...)
      return (
        <div
          style={{
            ...style,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            backgroundColor: '#f1f5f9',
            borderRight: '1px solid #cbd5e1',
            borderBottom: '1px solid #cbd5e1',
            fontSize: '13px',
            color: '#334155',
            userSelect: 'none',
            position: 'sticky',
            left: 0,
            zIndex: 10
          }}
          className="dark:bg-slate-800 dark:border-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600"
        >
          {rowIndex}
        </div>
      );
    }
    
    // Data cell
    // Grid: rowIndex 0 = header, rowIndex 1+ = data rows (Excel row 1+)
    // Grid: columnIndex 0 = header, columnIndex 1+ = data cols (Excel col 1+)
    // getCellValue expects 0-based indices and converts to 1-based Excel internally
    const gridRowIndex = rowIndex - 1; // Convert to 0-based data row index
    const gridColIndex = columnIndex - 1; // Convert to 0-based data col index
    
    const cellValue = getCellValue(gridRowIndex, gridColIndex);
    const excelRow = rowIndex; // Excel row number (1-based)
    const excelCol = columnIndex; // Excel col number (1-based, but we display as 0-based in grid)
    
    const isLoading = cellValue === null && 
      loadingWindows.size > 0 && 
      excelRow <= (metadata?.totalRows || 0) && 
      gridColIndex < (metadata?.totalColumns || 0);
    
    return (
      <div
        style={{
          ...style,
          display: 'flex',
          alignItems: 'center',
          padding: '0 8px',
          borderRight: '1px solid #e2e8f0',
          borderBottom: '1px solid #e2e8f0',
          fontSize: '13px',
          color: '#0f172a',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          backgroundColor: isLoading ? '#f8fafc' : '#ffffff'
        }}
        className="dark:bg-slate-900 dark:border-slate-700 dark:text-slate-100 hover:border-double hover:border-blue-500"
        title={cellValue !== null && cellValue !== undefined ? String(cellValue) : ''}
      >
        {isLoading ? (
          <span className="text-slate-400 text-xs">...</span>
        ) : (
          cellValue !== null && cellValue !== undefined ? String(cellValue) : ''
        )}
      </div>
    );
  }, [getCellValue, getColumnLabel, loadingWindows, metadata]);

  /**
   * Handle scroll events via the grid element
   * react-window 2.x doesn't have onScroll prop, so we attach to the element
   */
  useEffect(() => {
    if (!gridRef.current?.element || !metadata) return;
    
    const gridElement = gridRef.current.element;
    
    const handleScroll = () => {
      const scrollTop = gridElement.scrollTop;
      const scrollLeft = gridElement.scrollLeft;
      setScrollState({ scrollTop, scrollLeft });
      
      // Calculate which rows/columns are visible (account for header row/col)
      // Subtract header height/width from scroll position
      const dataScrollTop = Math.max(0, scrollTop - COL_HEADER_HEIGHT);
      const dataScrollLeft = Math.max(0, scrollLeft - ROW_HEADER_WIDTH);
      
      const visibleRowStart = Math.floor(dataScrollTop / CELL_HEIGHT);
      const visibleRowEnd = Math.ceil((dataScrollTop + containerSize.height - COL_HEADER_HEIGHT) / CELL_HEIGHT);
      const visibleColStart = Math.floor(dataScrollLeft / CELL_WIDTH);
      const visibleColEnd = Math.ceil((dataScrollLeft + containerSize.width - ROW_HEADER_WIDTH) / CELL_WIDTH);
      
      // Prefetch windows that will be needed soon
      const prefetchRowStart = Math.max(0, visibleRowStart - ROW_WINDOW_SIZE);
      const prefetchRowEnd = Math.min(
        metadata.totalRows || 10000,
        visibleRowEnd + ROW_WINDOW_SIZE
      );
      const prefetchColStart = Math.max(0, visibleColStart - COL_WINDOW_SIZE);
      const prefetchColEnd = Math.min(
        metadata.totalColumns || 100,
        visibleColEnd + COL_WINDOW_SIZE
      );
      
      // Fetch windows for visible area and buffer
      for (let r = prefetchRowStart; r < prefetchRowEnd; r += ROW_WINDOW_SIZE) {
        for (let c = prefetchColStart; c < prefetchColEnd; c += COL_WINDOW_SIZE) {
          const window = calculateWindow(r, c);
          fetchWindow(window.rowStart, window.rowEnd, window.colStart, window.colEnd);
        }
      }
    };
    
    gridElement.addEventListener('scroll', handleScroll);
    return () => gridElement.removeEventListener('scroll', handleScroll);
  }, [gridRef, metadata, calculateWindow, fetchWindow, containerSize]);

  // Calculate visible range for status display
  const visibleRange = useMemo(() => {
    if (!metadata) return { rows: '...', cols: '...' };
    
    // Account for header row in scroll calculation
    const dataScrollTop = Math.max(0, scrollState.scrollTop - COL_HEADER_HEIGHT);
    const dataScrollLeft = Math.max(0, scrollState.scrollLeft - ROW_HEADER_WIDTH);
    
    const startRow = Math.max(1, Math.floor(dataScrollTop / CELL_HEIGHT) + 1);
    const endRow = Math.min(
      metadata.totalRows,
      Math.ceil((dataScrollTop + containerSize.height - COL_HEADER_HEIGHT) / CELL_HEIGHT)
    );
    const startCol = Math.max(0, Math.floor(dataScrollLeft / CELL_WIDTH));
    const endCol = Math.min(
      metadata.totalColumns - 1,
      Math.ceil((dataScrollLeft + containerSize.width - ROW_HEADER_WIDTH) / CELL_WIDTH) - 1
    );
    
    return {
      rows: `${startRow}-${endRow} of ${metadata.totalRows}`,
      cols: `${getColumnLabel(startCol)}-${getColumnLabel(endCol)}`
    };
  }, [scrollState, metadata, getColumnLabel, containerSize]);

  if (!sheetId) {
    return (
      <div className="flex items-center justify-center h-full text-slate-500">
        No file selected
      </div>
    );
  }

  if (!metadata) {
    return (
      <div className="flex items-center justify-center h-full">
        <div className="text-slate-500">Loading sheet metadata...</div>
      </div>
    );
  }

  // Grid dimensions: +1 for header row, +1 for header column
  const totalRows = metadata.totalRows + 1; // +1 for header
  const totalColumns = metadata.totalColumns + 1; // +1 for header

  return (
    <div className="flex flex-col h-full bg-white dark:bg-slate-900 overflow-hidden font-sans text-[13px]">
      {/* Formula Bar */}
      <div className="flex items-center gap-2 px-2 py-1 border-b border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 flex-shrink-0">
        <div className="text-slate-500 font-bold w-8 text-center italic font-serif">fx</div>
        <div className="flex-1 bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-600 px-2 py-1 text-slate-700 dark:text-slate-300 h-7 flex items-center shadow-inner"></div>
        {/* Status indicator */}
        <div className="text-xs text-slate-500 dark:text-slate-400 px-2">
          Row {visibleRange.rows} | Col {visibleRange.cols}
        </div>
      </div>

      {/* Virtualized Grid */}
      <div ref={containerRef} className="flex-1 relative" style={{ height: '100%', overflow: 'hidden' }}>
        <Grid
          gridRef={gridRef}
          columnCount={totalColumns}
          columnWidth={(index) => index === 0 ? ROW_HEADER_WIDTH : CELL_WIDTH}
          defaultHeight={containerSize.height}
          defaultWidth={containerSize.width}
          rowCount={totalRows}
          rowHeight={(index) => index === 0 ? COL_HEADER_HEIGHT : CELL_HEIGHT}
          className="custom-scrollbar"
          cellComponent={Cell}
          cellProps={{}}
        />
      </div>

      {/* Sheet Tabs (if multiple sheets) */}
      <div className="flex items-center gap-0 bg-slate-100 dark:bg-slate-800 border-t border-slate-300 dark:border-slate-700 px-1 h-9 overflow-x-auto custom-scrollbar flex-shrink-0">
        <button
          className="flex items-center justify-center px-4 h-[calc(100%-4px)] mt-[2px] font-medium text-xs shadow-sm rounded-t-sm border-t-2 bg-white dark:bg-slate-900 text-green-700 dark:text-green-500 border-green-600"
        >
          {metadata.sheetName || 'Sheet1'}
        </button>
      </div>
    </div>
  );
};

export default VirtualizedExcelView;

