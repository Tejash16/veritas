// src/components/Preview/EnhancedDocumentPreview.jsx
import React, { useState, useEffect, useRef } from 'react';
import { ZoomIn, ZoomOut, RotateCw, Edit3, Save, X } from 'lucide-react';

const EnhancedDocumentPreview = ({ 
  documentData, 
  extractedValues, 
  onValueEdit, 
  onValueSelect,
  selectedValueId 
}) => {
  const [currentPage, setCurrentPage] = useState(0);
  const [zoom, setZoom] = useState(1.0);
  const [editingValue, setEditingValue] = useState(null);
  const [hoveredValue, setHoveredValue] = useState(null);
  const canvasRef = useRef(null);
  const [canvasContext, setCanvasContext] = useState(null);

  useEffect(() => {
    if (canvasRef.current) {
      const ctx = canvasRef.current.getContext('2d');
      setCanvasContext(ctx);
      drawPage();
    }
  }, [currentPage, zoom, extractedValues, selectedValueId, hoveredValue]);

  const drawPage = async () => {
    if (!canvasContext || !documentData?.page_images?.[currentPage]) return;

    const canvas = canvasRef.current;
    const img = new Image();
    
    img.onload = () => {
      // Clear canvas
      canvasContext.clearRect(0, 0, canvas.width, canvas.height);
      
      // Draw page image
      const scaledWidth = img.width * zoom;
      const scaledHeight = img.height * zoom;
      canvas.width = scaledWidth;
      canvas.height = scaledHeight;
      
      canvasContext.drawImage(img, 0, 0, scaledWidth, scaledHeight);
      
      // Draw value overlays
      drawValueOverlays();
    };
    
    img.src = `data:image/png;base64,${documentData.page_images[currentPage]}`;
  };

  const drawValueOverlays = () => {
    if (!extractedValues || !canvasContext) return;

    const pageValues = extractedValues.filter(v => v.page_number === currentPage + 1);
    
    pageValues.forEach(value => {
      const bbox = value.location?.bounding_box;
      if (!bbox || bbox.length !== 4) return;

      const [x1, y1, x2, y2] = bbox;
      const canvas = canvasRef.current;
      
      // Convert normalized coordinates to canvas coordinates
      const canvasX1 = x1 * canvas.width;
      const canvasY1 = y1 * canvas.height;
      const canvasX2 = x2 * canvas.width;
      const canvasY2 = y2 * canvas.height;
      
      // Determine overlay style based on state
      let overlayColor = getConfidenceColor(value.confidence);
      let strokeWidth = 2;
      
      if (selectedValueId === value.id) {
        overlayColor = '#3B82F6';
        strokeWidth = 3;
      } else if (hoveredValue === value.id) {
        overlayColor = '#8B5CF6';
        strokeWidth = 3;
      }
      
      // Draw bounding box
      canvasContext.strokeStyle = overlayColor;
      canvasContext.lineWidth = strokeWidth;
      canvasContext.setLineDash(selectedValueId === value.id ? [] : [5, 5]);
      canvasContext.strokeRect(canvasX1, canvasY1, canvasX2 - canvasX1, canvasY2 - canvasY1);
      
      // Draw confidence indicator
      const confidenceHeight = 20;
      canvasContext.fillStyle = overlayColor;
      canvasContext.fillRect(canvasX1, canvasY1 - confidenceHeight, (canvasX2 - canvasX1) * value.confidence, confidenceHeight);
      
      // Draw value label
      canvasContext.fillStyle = '#FFFFFF';
      canvasContext.font = '12px Arial';
      canvasContext.fillText(
        value.value, 
        canvasX1 + 5, 
        canvasY1 - 5
      );
    });
  };

  const getConfidenceColor = (confidence) => {
    if (confidence >= 0.8) return '#10B981'; // Green
    if (confidence >= 0.6) return '#F59E0B'; // Yellow
    return '#EF4444'; // Red
  };

  const handleCanvasClick = (event) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / canvas.width;
    const y = (event.clientY - rect.top) / canvas.height;
    
    // Find clicked value
    const pageValues = extractedValues.filter(v => v.page_number === currentPage + 1);
    const clickedValue = pageValues.find(value => {
      const bbox = value.location?.bounding_box;
      if (!bbox) return false;
      const [x1, y1, x2, y2] = bbox;
      return x >= x1 && x <= x2 && y >= y1 && y <= y2;
    });
    
    if (clickedValue) {
      onValueSelect(clickedValue.id);
    }
  };

  const handleValueEdit = (valueId, newData) => {
    onValueEdit(valueId, newData);
    setEditingValue(null);
  };

  return (
    <div className="bg-white border rounded-lg shadow-sm">
      {/* Header Controls */}
      <div className="flex items-center justify-between p-4 border-b">
        <div className="flex items-center space-x-4">
          <h3 className="text-lg font-medium">Document Preview</h3>
          <span className="text-sm text-gray-500">
            Page {currentPage + 1} of {documentData?.page_images?.length || 0}
          </span>
        </div>
        
        <div className="flex items-center space-x-2">
          <button
            onClick={() => setZoom(Math.max(0.5, zoom - 0.25))}
            className="p-1 text-gray-600 hover:text-gray-900"
          >
            <ZoomOut className="h-4 w-4" />
          </button>
          <span className="text-sm text-gray-600">{Math.round(zoom * 100)}%</span>
          <button
            onClick={() => setZoom(Math.min(2.0, zoom + 0.25))}
            className="p-1 text-gray-600 hover:text-gray-900"
          >
            <ZoomIn className="h-4 w-4" />
          </button>
        </div>
      </div>

      {/* Document Canvas */}
      <div className="relative bg-gray-100 min-h-96 overflow-auto">
        <canvas
          ref={canvasRef}
          onClick={handleCanvasClick}
          onMouseMove={(e) => {
            // Handle hover effects for values
            const canvas = canvasRef.current;
            const rect = canvas.getBoundingClientRect();
            const x = (e.clientX - rect.left) / canvas.width;
            const y = (e.clientY - rect.top) / canvas.height;
            
            const pageValues = extractedValues.filter(v => v.page_number === currentPage + 1);
            const hoveredVal = pageValues.find(value => {
              const bbox = value.location?.bounding_box;
              if (!bbox) return false;
              const [x1, y1, x2, y2] = bbox;
              return x >= x1 && x <= x2 && y >= y1 && y <= y2;
            });
            
            setHoveredValue(hoveredVal?.id || null);
          }}
          className="cursor-pointer"
          style={{ maxWidth: '100%', height: 'auto' }}
        />
      </div>

      {/* Page Navigation */}
      <div className="flex items-center justify-between p-4 border-t">
        <button
          onClick={() => setCurrentPage(Math.max(0, currentPage - 1))}
          disabled={currentPage === 0}
          className="px-3 py-1 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50"
        >
          Previous
        </button>
        
        <div className="flex items-center space-x-2">
          {Array.from({ length: documentData?.page_images?.length || 0 }, (_, i) => (
            <button
              key={i}
              onClick={() => setCurrentPage(i)}
              className={`w-8 h-8 text-xs rounded ${
                currentPage === i
                  ? 'bg-blue-600 text-white'
                  : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
              }`}
            >
              {i + 1}
            </button>
          ))}
        </div>
        
        <button
          onClick={() => setCurrentPage(Math.min((documentData?.page_images?.length || 1) - 1, currentPage + 1))}
          disabled={currentPage >= (documentData?.page_images?.length || 1) - 1}
          className="px-3 py-1 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50"
        >
          Next
        </button>
      </div>

      {/* Value Edit Modal */}
      {editingValue && (
        <ValueEditModal
          value={editingValue}
          onSave={handleValueEdit}
          onCancel={() => setEditingValue(null)}
        />
      )}
    </div>
  );
};

const ValueEditModal = ({ value, onSave, onCancel }) => {
  const [editedValue, setEditedValue] = useState(value.value);
  const [editedContext, setEditedContext] = useState(value.business_context);

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg p-6 max-w-md w-full mx-4">
        <h3 className="text-lg font-medium mb-4">Edit Extracted Value</h3>
        
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Value
            </label>
            <input
              type="text"
              value={editedValue}
              onChange={(e) => setEditedValue(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded-md"
            />
          </div>
          
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">
              Business Context
            </label>
            <textarea
              value={editedContext}
              onChange={(e) => setEditedContext(e.target.value)}
              rows={3}
              className="w-full px-3 py-2 border border-gray-300 rounded-md"
            />
          </div>
        </div>
        
        <div className="flex justify-end space-x-3 mt-6">
          <button
            onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50"
          >
            Cancel
          </button>
          <button
            onClick={() => onSave(value.id, { value: editedValue, business_context: editedContext })}
            className="px-4 py-2 text-sm font-medium text-white bg-blue-600 border border-transparent rounded-md hover:bg-blue-700"
          >
            Save Changes
          </button>
        </div>
      </div>
    </div>
  );
};

export default EnhancedDocumentPreview;