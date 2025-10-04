// src/components/Viewer/ExcelViewer.jsx
import React, { useEffect, useMemo, useRef, useState, useCallback } from "react";
import { FixedSizeGrid as Grid } from "react-window";
import { ZoomIn, ZoomOut } from "lucide-react";
import { getExcelMeta } from "../../services/api";
import { useExcelTiles } from "../../hooks/useExcelTiles";

function colToLetter(col) {
  let letter = '';
  while (col > 0) {
    const remainder = (col - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function a1ToRC(a1) {
  const match = /^([A-Z]+)(\d+)$/i.exec(String(a1).toUpperCase());
  if (!match) return null;
  
  const letters = match[1];
  let col = 0;
  for (const ch of letters) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  
  return { row: parseInt(match[2], 10), col };
}

export default function ExcelViewer({
  fileId,
  activeSheet,
  highlightedCells = [],
  onSheetChange,
  height = 600,
  width = 1000
}) {
  const [meta, setMeta] = useState([]);
  const [sheet, setSheet] = useState(activeSheet || "");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [zoom, setZoom] = useState(100);
  const [scrollPos, setScrollPos] = useState({ scrollLeft: 0, scrollTop: 0 });
  const [tilesLoaded, setTilesLoaded] = useState(false);
  
  const mainGridRef = useRef(null);

  const ROW_HEIGHT = Math.round(28 * (zoom / 100));
  const COL_WIDTH = Math.round(100 * (zoom / 100));
  const FONT_SIZE = Math.round(13 * (zoom / 100));
  const HEADER_SIZE = 30;
  const BUFFER_ROWS = 100;
  const BUFFER_COLS = 26;

  const dims = useMemo(() => {
    const s = meta.find((x) => x.name === sheet);
    return { 
      rows: (s?.rows || 1) + BUFFER_ROWS,
      cols: (s?.cols || 1) + BUFFER_COLS
    };
  }, [meta, sheet]);

  const { fetchTile, getCell, tileRows, tileCols, clearCache } = useExcelTiles(fileId);

  // Validate and parse highlighted cells
  const highlightedCellsSet = useMemo(() => {
    const set = new Set();
    
    highlightedCells.forEach(({ sheet: cellSheet, cell }) => {
      // Only process cells for current sheet
      if (cellSheet !== sheet) return;
      
      // Validate cell format
      const parsed = a1ToRC(cell);
      if (!parsed) {
        console.warn(`[ExcelViewer] Invalid cell reference: ${cell}`);
        return;
      }
      
      set.add(cell.toUpperCase());
    });
    
    console.log('[ExcelViewer] Valid highlighted cells:', Array.from(set));
    return set;
  }, [highlightedCells, sheet]);

  // Load metadata
  useEffect(() => {
    if (!fileId) return;
    
    setLoading(true);
    setError(null);
    
    getExcelMeta(fileId)
      .then((data) => {
        console.log('[ExcelViewer] Loaded meta:', data);
        setMeta(data.sheets || []);
        
        // Set initial sheet
        if (!sheet && data.sheets?.length > 0) {
          const firstSheet = activeSheet || data.sheets[0].name;
          setSheet(firstSheet);
        }
        
        setLoading(false);
      })
      .catch((err) => {
        console.error("[ExcelViewer] Failed to load Excel metadata:", err);
        setError("Failed to load Excel file");
        setLoading(false);
      });
  }, [fileId]);

  // Handle sheet changes - with validation
  useEffect(() => {
    if (!activeSheet) return;
    
    // Validate sheet exists
    const sheetExists = meta.some(s => s.name === activeSheet);
    if (!sheetExists) {
      console.warn(`[ExcelViewer] Sheet "${activeSheet}" not found in meta. Available:`, meta.map(s => s.name));
      return; // Don't change sheet if invalid
    }
    
    if (sheet !== activeSheet) {
      console.log('[ExcelViewer] Switching to sheet:', activeSheet);
      setSheet(activeSheet);
      setTilesLoaded(false);
      clearCache();
      onSheetChange?.(activeSheet);
    }
  }, [activeSheet, meta]);

  // Preload visible tiles
  useEffect(() => {
    if (!sheet || tilesLoaded || !meta.length) return;

    const preloadVisibleTiles = async () => {
      const gridHeight = height - HEADER_SIZE;
      const gridWidth = width - HEADER_SIZE;
      const visibleRows = Math.ceil(gridHeight / ROW_HEIGHT) + 5;
      const visibleCols = Math.ceil(gridWidth / COL_WIDTH) + 5;

      console.log('[ExcelViewer] Preloading tiles:', { sheet, visibleRows, visibleCols });

      const promises = [];
      for (let r = 0; r < Math.min(visibleRows, dims.rows); r += tileRows) {
        for (let c = 0; c < Math.min(visibleCols, dims.cols); c += tileCols) {
          promises.push(fetchTile(sheet, r + 1, c + 1).catch(err => {
            console.warn(`Failed to fetch tile at ${r},${c}:`, err);
            return null;
          }));
        }
      }

      try {
        await Promise.allSettled(promises);
        console.log('[ExcelViewer] Preload complete');
        setTilesLoaded(true);
      } catch (err) {
        console.error('[ExcelViewer] Preload failed:', err);
        setTilesLoaded(true); // Set anyway to prevent infinite loading
      }
    };

    // Add small delay to ensure sheet change is complete
    const timer = setTimeout(preloadVisibleTiles, 100);
    return () => clearTimeout(timer);
  }, [sheet, height, width, ROW_HEIGHT, COL_WIDTH, meta, tilesLoaded]);

  // Scroll to highlighted cell - with validation
  useEffect(() => {
    if (highlightedCells.length === 0 || !mainGridRef.current || !tilesLoaded) return;
    
    const cellToScroll = highlightedCells.find(c => c.sheet === sheet);
    if (!cellToScroll) return;

    const parsed = a1ToRC(cellToScroll.cell);
    if (!parsed) {
      console.warn(`[ExcelViewer] Cannot scroll to invalid cell: ${cellToScroll.cell}`);
      return;
    }

    const { row, col } = parsed;
    const rowIdx = row - 1;
    const colIdx = col - 1;

    // Validate bounds
    if (rowIdx < 0 || rowIdx >= dims.rows || colIdx < 0 || colIdx >= dims.cols) {
      console.warn(`[ExcelViewer] Cell ${cellToScroll.cell} is out of bounds`);
      return;
    }

    console.log('[ExcelViewer] Scrolling to:', cellToScroll, { row, col });

    setTimeout(() => {
      if (mainGridRef.current) {
        try {
          const outer = mainGridRef.current._outerRef;
          const targetTop = rowIdx * ROW_HEIGHT - Math.max(0, (outer.clientHeight) / 2 - ROW_HEIGHT / 2);
          const targetLeft = colIdx * COL_WIDTH - Math.max(0, (outer.clientWidth) / 2 - COL_WIDTH / 2);
          
          mainGridRef.current.scrollTo({
            scrollTop: Math.max(0, targetTop),
            scrollLeft: Math.max(0, targetLeft)
          });
        } catch (err) {
          console.error('[ExcelViewer] Scroll failed:', err);
        }
      }
    }, 200);
  }, [highlightedCells, sheet, zoom, tilesLoaded]);

  const handleScroll = useCallback(({ scrollLeft, scrollTop }) => {
    setScrollPos({ scrollLeft, scrollTop });
  }, []);

  // Cell renderer with error handling
  const Cell = useCallback(({ columnIndex, rowIndex, style }) => {
    const r = rowIndex + 1;
    const c = columnIndex + 1;
    const cellRef = `${colToLetter(c)}${r}`;

    // useEffect(() => {
    //   const r0 = Math.floor(rowIndex / tileRows) * tileRows + 1;
    //   const c0 = Math.floor(columnIndex / tileCols) * tileCols + 1;
    //   fetchTile(sheet, r0, c0).catch(() => {}); // Silently fail
    // }, [rowIndex, columnIndex]);

    const val = getCell(sheet, r, c);
    const isHighlighted = highlightedCellsSet.has(cellRef);

    const isNumber = val != null && !isNaN(val) && val !== '';
    const cellClass = isNumber ? 'justify-end' : 'justify-start';

    return (
      <div
        style={{
          ...style,
          fontSize: `${FONT_SIZE}px`,
          lineHeight: `${ROW_HEIGHT - 4}px`
        }}
        className={`border-r border-b border-gray-300 flex items-center px-2 ${cellClass} ${
          isHighlighted
            ? "bg-yellow-200 ring-2 ring-inset ring-blue-600 font-bold"
            : "bg-white hover:bg-gray-50"
        }`}
        title={val != null ? String(val) : ""}
      >
        <div className="truncate w-full">
          {val != null ? String(val) : ""}
        </div>
      </div>
    );
  }, [sheet, FONT_SIZE, ROW_HEIGHT, highlightedCellsSet, getCell]);

  const handleZoomIn = () => setZoom(Math.min(200, zoom + 10));
  const handleZoomOut = () => setZoom(Math.max(50, zoom - 10));
  const handleZoomReset = () => setZoom(100);

  if (loading) {
    return (
      <div className="flex flex-col border rounded-xl overflow-hidden bg-white h-full">
        <div className="flex items-center justify-center flex-1">
          <div className="text-gray-500">Loading Excel data...</div>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex flex-col border rounded-xl overflow-hidden bg-white h-full">
        <div className="flex items-center justify-center flex-1">
          <div className="text-red-500">{error}</div>
        </div>
      </div>
    );
  }

  const gridHeight = height - HEADER_SIZE;
  const gridWidth = width - HEADER_SIZE;

  return (
    <div className="flex flex-col border rounded-xl overflow-hidden bg-white shadow-sm">
      {/* Header - same as before */}
      <div className="flex items-center justify-between gap-2 px-3 py-2 border-b bg-gray-50">
        <div className="flex items-center gap-2">
          <div className="text-sm font-medium text-gray-700">Excel</div>
          <div className="flex gap-1 overflow-x-auto">
            {meta.map((s) => (
              <button
                key={s.name}
                onClick={() => {
                  setSheet(s.name);
                  setTilesLoaded(false);
                  clearCache();
                  onSheetChange?.(s.name);
                }}
                className={`px-3 py-1 rounded text-xs whitespace-nowrap transition ${
                  s.name === sheet
                    ? "bg-white text-gray-900 font-medium shadow-sm border border-gray-300"
                    : "text-gray-600 hover:bg-gray-100"
                }`}
              >
                {s.name}
              </button>
            ))}
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="text-xs text-gray-500 font-mono">
            {highlightedCells.length > 0 
              ? highlightedCells.map(c => `${c.cell}`).join(', ')
              : "No cells highlighted"}
          </div>

          <div className="flex items-center gap-1 border-l pl-3">
            <button onClick={handleZoomOut} className="p-1.5 hover:bg-gray-200 rounded transition">
              <ZoomOut className="h-4 w-4 text-gray-600" />
            </button>
            <button onClick={handleZoomReset} className="px-2 py-1 hover:bg-gray-200 rounded text-xs font-medium text-gray-700 transition min-w-[45px]">
              {zoom}%
            </button>
            <button onClick={handleZoomIn} className="p-1.5 hover:bg-gray-200 rounded transition">
              <ZoomIn className="h-4 w-4 text-gray-600" />
            </button>
          </div>
        </div>
      </div>

      {/* Grid - same structure as before */}
      <div className="relative bg-gray-200" style={{ height, width, overflow: 'hidden' }}>
        {/* Headers and grid code same as before */}
        <div style={{ position: 'absolute', width: HEADER_SIZE, height: HEADER_SIZE, top: 0, left: 0, zIndex: 30 }} className="bg-gray-100 border-r border-b border-gray-300" />
        
        <div style={{ position: 'absolute', top: 0, left: HEADER_SIZE, width: gridWidth, height: HEADER_SIZE, zIndex: 20, overflow: 'hidden', backgroundColor: '#f3f4f6' }}>
          <div style={{ transform: `translateX(-${scrollPos.scrollLeft}px)`, width: dims.cols * COL_WIDTH }}>
            <div style={{ display: 'flex' }}>
              {Array.from({ length: dims.cols }, (_, i) => (
                <div key={i} style={{ width: COL_WIDTH, height: HEADER_SIZE, fontSize: '11px' }} className="flex items-center justify-center bg-gray-100 border-r border-b border-gray-300 font-semibold text-gray-700 flex-shrink-0">
                  {colToLetter(i + 1)}
                </div>
              ))}
            </div>
          </div>
        </div>

        <div style={{ position: 'absolute', top: HEADER_SIZE, left: 0, width: HEADER_SIZE, height: gridHeight, zIndex: 20, overflow: 'hidden', backgroundColor: '#f3f4f6' }}>
          <div style={{ transform: `translateY(-${scrollPos.scrollTop}px)`, height: dims.rows * ROW_HEIGHT }}>
            {Array.from({ length: dims.rows }, (_, i) => (
              <div key={i} style={{ width: HEADER_SIZE, height: ROW_HEIGHT, fontSize: '11px' }} className="flex items-center justify-center bg-gray-100 border-r border-b border-gray-300 font-semibold text-gray-700">
                {i + 1}
              </div>
            ))}
          </div>
        </div>

        <div style={{ position: 'absolute', top: HEADER_SIZE, left: HEADER_SIZE, width: gridWidth, height: gridHeight, zIndex: 10 }}>
          <Grid
            ref={mainGridRef}
            columnCount={dims.cols}
            columnWidth={COL_WIDTH}
            height={gridHeight}
            rowCount={dims.rows}
            rowHeight={ROW_HEIGHT}
            width={gridWidth}
            onScroll={handleScroll}
            overscanRowCount={5}
            overscanColumnCount={5}
            itemKey={({ columnIndex, rowIndex }) => `${sheet}-${rowIndex}-${columnIndex}`}
          >
            {Cell}
          </Grid>
        </div>
      </div>

      <div className="px-3 py-1.5 border-t bg-gray-50 flex items-center justify-between text-xs text-gray-600">
        <div>Sheet: {sheet}</div>
        <div className="flex items-center gap-3">
          <span>Zoom: {zoom}%</span>
          {highlightedCells.length > 0 && <span>{highlightedCells.length} cell(s) highlighted</span>}
        </div>
      </div>
    </div>
  );
}