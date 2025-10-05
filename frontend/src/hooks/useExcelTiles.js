// src/hooks/useExcelTiles.js
import { useRef } from "react";
import { getExcelPage } from "../services/api";

/**
 * Hook for managing Excel tile cache with server-side pagination.
 * Fetches tiles on demand and caches them for performance.
 * 
 * @param {string} fileId - The Excel file identifier
 * @param {number} tileRows - Number of rows per tile (default: 200)
 * @param {number} tileCols - Number of columns per tile (default: 50)
 */
export function useExcelTiles(fileId, tileRows = 1000, tileCols = 250) {
  const cache = useRef(new Map());

  /**
   * Fetch a tile from the server or return from cache
   * @param {string} sheet - Sheet name
   * @param {number} r0 - Starting row (1-based)
   * @param {number} c0 - Starting column (1-based)
   */
  async function fetchTile(sheet, r0, c0) {
    const key = `${sheet}:${r0}:${c0}`;
    
    // Return from cache if available
    if (cache.current.has(key)) {
      return cache.current.get(key);
    }

    // Fetch from server
    const page = await getExcelPage({
      fileId,
      sheet,
      r0,
      r1: r0 + tileRows - 1,
      c0,
      c1: c0 + tileCols - 1
    });

    // Cache the result
    cache.current.set(key, page);
    return page;
  }

  /**
   * Get a cell value from the cache
   * @param {string} sheet - Sheet name
   * @param {number} row - Row number (1-based)
   * @param {number} col - Column number (1-based)
   */
  function getCell(sheet, row, col) {
    for (const page of cache.current.values()) {
      if (page.sheet !== sheet) continue;
      
      if (page.r0 <= row && row <= page.r1 && page.c0 <= col && col <= page.c1) {
        return page.data[row - page.r0][col - page.c0];
      }
    }
    return null;
  }

  /**
   * Clear the cache (useful when switching files)
   */
  function clearCache() {
    cache.current.clear();
  }

  return {
    fetchTile,
    getCell,
    clearCache,
    tileRows,
    tileCols,
    cache: cache.current
  };
}