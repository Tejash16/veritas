// src/components/Dashboard/Preview.jsx
import React, { useState, useEffect } from 'react';
import ExcelViewer from '../Viewer/ExcelViewer';
import { enhancedApiService } from '../../services/api';

const Preview = ({ preview, sessionId }) => {
  const [active, setActive] = useState(null);
  const [excelFileId, setExcelFileId] = useState(null);
  const [uploadSessionId, setUploadSessionId] = useState(null);
  const [availableSheets, setAvailableSheets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  /**
   * Parse multiple cell references separated by "!"
   * Example: "O10!B5!C3" -> ["O10", "B5", "C3"]
   * Single cell: "O10" -> ["O10"]
   */
  const parseMultipleCells = (cellRef) => {
    if (!cellRef || typeof cellRef !== 'string') return [];
    
    // Split by "!" to get multiple cells
    const cells = cellRef.split('!').map(c => c.trim()).filter(c => c.length > 0);
    
    // Validate each cell format (e.g., A1, B10, AA100)
    const cellPattern = /^[A-Z]+\d+$/i;
    const validCells = cells.filter(cell => cellPattern.test(cell));
    
    if (validCells.length !== cells.length) {
      const invalid = cells.filter(cell => !cellPattern.test(cell));
      console.warn('[Preview] Invalid cell references filtered out:', invalid, 'from:', cellRef);
    }
    
    return validCells;
  };

  /**
   * Validate if a sheet name exists in available sheets
   */
  const isValidSheet = (sheetName) => {
    if (!sheetName) return false;
    return availableSheets.some(s => 
      s.toLowerCase() === sheetName.toLowerCase()
    );
  };

  // Transform detailed_results into values array for the viewer
  const transformedValues = React.useMemo(() => {
    if (!preview || !Array.isArray(preview)) {
      console.log('[Preview] No preview data or not array:', preview);
      return [];
    }

    console.log('[Preview] Raw preview data:', preview);
    console.log('[Preview] Total items in preview:', preview.length);

    // Count statuses in raw data
    const statusCounts = {
      matched: 0,
      mismatched: 0,
      unverifiable: 0,
      other: 0
    };
    
    preview.forEach(item => {
      const status = item.validation_status?.toLowerCase();
      if (status === 'matched') statusCounts.matched++;
      else if (status === 'mismatched') statusCounts.mismatched++;
      else if (status === 'unverifiable') statusCounts.unverifiable++;
      else statusCounts.other++;
    });
    
    console.log('[Preview] Status distribution:', statusCounts);

    const values = preview
      // FILTER: Only include matched and mismatched, exclude unverifiable
      .filter(result => {
        const status = result.validation_status?.toLowerCase();
        const shouldInclude = status === 'matched' || status === 'mismatched';
        
        if (!shouldInclude) {
          console.log(`[Preview] Filtering out ${result.pdf_value_id} with status: ${status}`);
        }
        
        return shouldInclude;
      })
      .map((result, index) => {
        // Check if excel_match exists and has data
        const excelMatch = result.excel_match || {};
        
        // Skip if no excel_match or it's empty
        if (!excelMatch.Source_Sheet && !excelMatch.Excel_Cell_used_for_match) {
          console.log(`[Preview] Skipping value ${result.pdf_value_id}: no excel match`);
          return null;
        }

        const sheetName = excelMatch.Source_Sheet || '';
        const rawCellRef = excelMatch.Excel_Cell_used_for_match || '';
        
        // Parse multiple cells (e.g., "O10!B5!C3" -> ["O10", "B5", "C3"])
        const cellRefs = parseMultipleCells(rawCellRef);

        // Validate parsed data
        if (!sheetName || cellRefs.length === 0) {
          console.log(`[Preview] Skipping value ${result.pdf_value_id}: invalid sheet or cells`, {
            sheetName,
            rawCellRef,
            parsedCells: cellRefs
          });
          return null;
        }

        return {
          id: result.pdf_value_id || `value_${index}`,
          display: result.pdf_value || '',
          status: result.validation_status || 'unverified',
          confidence: result.confidence || 0,
          page: result.page || 1,
          sheet: sheetName,
          cells: cellRefs, // Array of cell references ["O10", "B5"]
          cellsDisplay: cellRefs.join(', '), // "O10, B5" for display
          rawCellRef: rawCellRef, // Keep original for debugging
          context: result.pdf_context || '',
          auditReasoning: result.audit_reasoning || '',
          excelValue: excelMatch.excel_value || '',
          calculationBasis: excelMatch.calculation_basis || null,
          matchSource: excelMatch.match_source || null,
          matchConfidence: excelMatch.match_confidence || 0
        };
      })
      .filter(v => v !== null); // Remove null entries

    // Count final statuses
    const finalStatusCounts = {
      matched: values.filter(v => v.status === 'matched').length,
      mismatched: values.filter(v => v.status === 'mismatched').length
    };

    console.log('[Preview] Transformed values:', values);
    console.log('[Preview] Final counts - Matched:', finalStatusCounts.matched, 'Mismatched:', finalStatusCounts.mismatched);
    console.log('[Preview] Total values shown:', values.length, '(filtered out', statusCounts.unverifiable, 'unverifiable)');
    
    // Log multi-cell values specifically
    const multiCellValues = values.filter(v => v.cells.length > 1);
    if (multiCellValues.length > 0) {
      console.log('[Preview] Values with multiple cells:', multiCellValues.map(v => ({
        id: v.id,
        value: v.display,
        cells: v.cells
      })));
    }
    
    return values;
  }, [preview, availableSheets]);

  // Load upload session ID and Excel file ID
  useEffect(() => {
    const loadExcelData = async () => {
      if (!sessionId) {
        setError('No session ID provided');
        setLoading(false);
        return;
      }

      try {
        setLoading(true);
        setError(null);

        console.log('[Preview] Loading data for audit session:', sessionId);

        const auditResults = await enhancedApiService.getAuditResults(sessionId);
        console.log('[Preview] Audit results:', auditResults);

        const uploadSessionIdFromAudit = auditResults.upload_session_id;
        if (!uploadSessionIdFromAudit) {
          throw new Error('Upload session ID not found in audit results');
        }

        setUploadSessionId(uploadSessionIdFromAudit);

        const validationData = await enhancedApiService.getValidationData(uploadSessionIdFromAudit);
        console.log('[Preview] Validation data:', validationData);
        
        // Extract available sheet names for validation
        if (validationData.excel_documents || validationData.excel_file_id) {
          try {
            const fileId = validationData.excel_file_id || validationData.excel_documents[0]?.file_id;
            if (fileId) {
              const meta = await enhancedApiService.getExcelMeta(fileId);
              const sheets = meta.sheets?.map(s => s.name) || [];
              setAvailableSheets(sheets);
              console.log('[Preview] Available sheets:', sheets);
            }
          } catch (err) {
            console.warn('[Preview] Could not load sheet names:', err);
          }
        }
        
        if (validationData.excel_file_id) {
          setExcelFileId(validationData.excel_file_id);
        } else if (validationData.excel_documents?.[0]?.file_id) {
          setExcelFileId(validationData.excel_documents[0].file_id);
        } else {
          setError('No Excel file found for this session');
        }

        setLoading(false);
      } catch (err) {
        console.error('[Preview] Failed to load Excel data:', err);
        setError(`Failed to load Excel file: ${err.message}`);
        setLoading(false);
      }
    };

    loadExcelData();
  }, [sessionId]);

  const getStatusClass = (status) => {
    const statusLower = status?.toLowerCase();
    switch (statusLower) {
      case 'matched':
        return 'bg-green-100 text-green-700 border-green-300';
      case 'mismatched':
        return 'bg-red-100 text-red-700 border-red-300';
      case 'unverifiable':
        return 'bg-yellow-100 text-yellow-700 border-yellow-300';
      default:
        return 'bg-gray-100 text-gray-700 border-gray-300';
    }
  };

  const handleValueClick = (value) => {
    console.log('[Preview] Value clicked:', {
      id: value.id,
      display: value.display,
      sheet: value.sheet,
      cells: value.cells,
      cellCount: value.cells.length,
      status: value.status
    });
    
    // Validate sheet
    const isSheetValid = isValidSheet(value.sheet);
    if (!isSheetValid) {
      console.warn(`[Preview] Invalid sheet name: "${value.sheet}". Available sheets:`, availableSheets);
      return;
    }
    
    // Validate all cells
    const cellPattern = /^[A-Z]+\d+$/i;
    const allCellsValid = value.cells.every(cell => cellPattern.test(cell));
    if (!allCellsValid) {
      const invalidCells = value.cells.filter(cell => !cellPattern.test(cell));
      console.warn(`[Preview] Invalid cell references:`, invalidCells);
      return;
    }
    
    setActive(value);
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto"></div>
          <p className="mt-4 text-gray-600">Loading preview...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-center">
          <p className="text-red-600 font-medium">{error}</p>
          <p className="text-sm text-gray-500 mt-2">Please try refreshing the page</p>
        </div>
      </div>
    );
  }

  if (transformedValues.length === 0) {
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-center">
          <p className="text-gray-600">No matched or mismatched values found</p>
          <p className="text-sm text-gray-500 mt-2">
            Only showing values with status "matched" or "mismatched"
          </p>
          <p className="text-xs text-gray-400 mt-2">
            Check console for details about filtered values
          </p>
        </div>
      </div>
    );
  }

  // Calculate status counts for display
  const matchedCount = transformedValues.filter(v => v.status === 'matched').length;
  const mismatchedCount = transformedValues.filter(v => v.status === 'mismatched').length;

  return (
    <div className="grid grid-cols-[400px,1fr] gap-4 h-[calc(100vh-400px)] min-h-[600px]">
      {/* LEFT PANEL: PDF Values */}
      <div className="border rounded-lg overflow-hidden flex flex-col bg-white shadow-sm">
        <div className="px-4 py-3 border-b bg-gray-50">
          <h3 className="text-sm font-semibold text-gray-900">
            PDF Values ({transformedValues.length})
          </h3>
          <div className="flex gap-3 text-xs text-gray-600 mt-1">
            <span className="flex items-center gap-1">
              <span className="w-2 h-2 rounded-full bg-green-500"></span>
              Matched: {matchedCount}
            </span>
            <span className="flex items-center gap-1">
              <span className="w-2 h-2 rounded-full bg-red-500"></span>
              Mismatched: {mismatchedCount}
            </span>
          </div>
          <p className="text-xs text-gray-500 mt-1">
            Click a value to highlight in Excel
          </p>
        </div>

        <div className="overflow-y-auto flex-1">
          {transformedValues.map((value) => {
            const isSheetValid = isValidSheet(value.sheet);
            const cellPattern = /^[A-Z]+\d+$/i;
            const allCellsValid = value.cells.every(cell => cellPattern.test(cell));
            const isClickable = isSheetValid && allCellsValid;
            
            return (
              <button
                key={value.id}
                onClick={() => handleValueClick(value)}
                disabled={!isClickable}
                className={`w-full text-left px-4 py-3 border-b transition-colors ${
                  !isClickable 
                    ? 'opacity-50 cursor-not-allowed bg-gray-50' 
                    : 'hover:bg-gray-50'
                } ${
                  active?.id === value.id
                    ? 'bg-blue-50 border-l-4 border-l-blue-600'
                    : 'border-l-4 border-l-transparent'
                }`}
              >
                {/* Value and Status */}
                <div className="flex justify-between items-start mb-2">
                  <div className="font-semibold text-lg text-gray-900 truncate pr-2">
                    {value.display}
                  </div>
                  <span
                    className={`text-xs px-2 py-1 rounded-full font-medium border whitespace-nowrap ${getStatusClass(
                      value.status
                    )}`}
                  >
                    {value.status.toUpperCase()}
                  </span>
                </div>

                {/* Context */}
                {value.context && (
                  <div className="text-xs text-gray-600 mb-2 line-clamp-2">
                    {value.context}
                  </div>
                )}

                {/* Metadata */}
                <div className="grid grid-cols-2 gap-2 text-xs text-gray-500">
                  <div>
                    <span className="font-medium">Confidence:</span>{' '}
                    {value.matchConfidence ? `${(value.matchConfidence * 100).toFixed(0)}%` : 'N/A'}
                  </div>
                  <div>
                    <span className="font-medium">Page:</span> {value.page}
                  </div>
                  <div className="col-span-2">
                    <span className="font-medium">Excel:</span>{' '}
                    {value.sheet}!{value.cellsDisplay}
                    {value.cells.length > 1 && (
                      <span className="ml-1 text-blue-600 font-semibold">
                        ({value.cells.length} cells)
                      </span>
                    )}
                    {!isSheetValid && (
                      <span className="ml-1 text-red-500">(Invalid sheet)</span>
                    )}
                    {!allCellsValid && (
                      <span className="ml-1 text-red-500">(Invalid cells)</span>
                    )}
                  </div>
                  {value.matchSource && (
                    <div className="col-span-2">
                      <span className="font-medium">Match Type:</span> {value.matchSource}
                    </div>
                  )}
                  {value.calculationBasis && (
                    <div className="col-span-2">
                      <span className="font-medium">Basis:</span> {value.calculationBasis}
                    </div>
                  )}
                </div>

                {/* Excel Value */}
                {value.excelValue && value.excelValue !== value.display && (
                  <div className="mt-2 pt-2 border-t border-gray-200">
                    <div className="text-xs text-gray-500">
                      <span className="font-medium">Excel Value:</span>{' '}
                      <span className="text-gray-700">{value.excelValue}</span>
                    </div>
                  </div>
                )}

                {/* Debug Info for invalid cells */}
                {!isClickable && (
                  <div className="mt-2 pt-2 border-t border-red-200 bg-red-50 -mx-4 px-4 py-2">
                    <div className="text-xs text-red-600">
                      <span className="font-medium">Cannot highlight:</span>
                      {!isSheetValid && <div>Sheet "{value.sheet}" not found</div>}
                      {!allCellsValid && (
                        <div>Invalid cell format: {value.rawCellRef}</div>
                      )}
                    </div>
                  </div>
                )}

                {/* Audit Reasoning */}
                {active?.id === value.id && value.auditReasoning && (
                  <div className="mt-2 pt-2 border-t border-gray-200">
                    <div className="text-xs text-gray-600 italic">
                      {value.auditReasoning}
                    </div>
                  </div>
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* RIGHT PANEL: Excel Viewer */}
      <div className="border rounded-lg overflow-hidden bg-white shadow-sm">
        {excelFileId && active ? (
          <ExcelViewer
            fileId={excelFileId}
            activeSheet={active.sheet}
            highlightedCells={active.cells.map(cell => ({ 
              sheet: active.sheet, 
              cell: cell 
            }))}
            height={600}
            width={window.innerWidth - 500}
          />
        ) : excelFileId ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center">
              <p className="text-gray-600">Select a value to view in Excel</p>
              <p className="text-sm text-gray-500 mt-2">
                Showing {matchedCount} matched and {mismatchedCount} mismatched values
              </p>
              {availableSheets.length > 0 && (
                <p className="text-xs text-gray-400 mt-2">
                  Available sheets: {availableSheets.join(', ')}
                </p>
              )}
            </div>
          </div>
        ) : (
          <div className="flex items-center justify-center h-full">
            <div className="text-center">
              <p className="text-gray-600">Excel file not available</p>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default Preview;