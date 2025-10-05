import React, { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import { useNavigate } from 'react-router-dom';
import { 
  Upload, CheckCircle, X, FileText, Table, Brain, 
  ArrowRight, Zap, Users
} from 'lucide-react';
import toast from 'react-hot-toast';

import { apiService } from '../../services/api';
import LoadingSpinner from '../Common/LoadingSpinner';

const EnhancedFileUpload = () => {
  const navigate = useNavigate();
  
  // Core state
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [sessionId, setSessionId] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [processing, setProcessing] = useState(false);

  // File upload handling
  const onDrop = useCallback(async (acceptedFiles) => {
    if (acceptedFiles.length === 0) return;

    // Validate file types
    const pdfFiles = acceptedFiles.filter(file => file.type === 'application/pdf');
    const excelFiles = acceptedFiles.filter(file => 
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel'
    );

    if (pdfFiles.length !== 1) {
      toast.error('Please upload exactly one PDF file');
      return;
    }

    if (excelFiles.length !== 1) {
      toast.error('Please upload exactly one Excel file');
      return;
    }

    await uploadFiles(acceptedFiles);
  }, []);

  const uploadFiles = async (files) => {
    setUploading(true);
    // console.log('[Upload] Starting enhanced upload...');

    try {
      const formData = new FormData();
      files.forEach(file => {
        formData.append('files', file);
      });

      const response = await apiService.uploadDocuments(formData);
      // console.log('[Upload] Upload completed:', response);

      setUploadedFiles(response.documents);
      setSessionId(response.session_id);
      

    } catch (error) {
      // console.error('[Upload] Upload failed:', error);
      toast.error('Upload failed: ' + error.message);
    } finally {
      setUploading(false);
    }
  };

  const startComprehensiveProcessing = async () => {
    if (!sessionId) {
      toast.error('No session found. Please upload files first.');
      return;
    }

    setProcessing(true);

    try {
      // console.log('[Processing] Starting comprehensive processing...');
      
      // Step 1: Process documents for extraction
      const processingResponse = await apiService.processDocuments(sessionId);
      // console.log('[Processing] Document processing completed:', processingResponse);


      // Step 2: Navigate directly to validation (skip mapping)
      // console.log('[Processing] Navigating to direct validation...');
      navigate(`/validation/${sessionId}`);

    } catch (error) {
      // console.error('[Processing] Processing failed:', error);
      toast.error('Processing failed: ' + error.message);
    } finally {
      setProcessing(false);
    }
  };

  const removeFile = (fileId) => {
    setUploadedFiles(prev => prev.filter(file => file.id !== fileId));
    if (uploadedFiles.length <= 1) {
      setSessionId(null);
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/pdf': ['.pdf'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    disabled: uploading || processing || uploadedFiles.length > 0
  });

  return (
    <>
      {/* Full Screen Loading Overlay */}
      {processing && (
        <div className="fixed inset-0 bg-gray-900 bg-opacity-75 flex items-center justify-center z-50">
          <div className="bg-white rounded-xl p-8 shadow-2xl flex flex-col items-center space-y-4">
            <LoadingSpinner size="xl" />
            <p className="text-2xl font-semibold text-gray-900">Extracting Values</p>
            <p className="text-sm text-gray-500">Please wait while we process your documents...</p>
          </div>
        </div>
      )}

      <div className="max-w-6xl mx-auto space-y-8">
        {/* Enhanced Header */}
        <div className="text-center">
          <div className="flex items-center justify-center mb-4">
            <Brain className="h-12 w-12 text-blue-600 mr-4" />
            <div>
              <h1 className="text-4xl font-bold text-gray-900">Veritas AI Auditor</h1>
              <p className="text-xl text-blue-600 font-medium">Direct Value Validation Edition</p>
            </div>
          </div>
          <p className="text-lg text-gray-600 max-w-3xl mx-auto">
            Advanced enterprise presentation validation with <strong>100% coverage</strong> - 
            validate ALL extracted values directly without mapping bottlenecks
          </p>
        </div>

        {/* Upload Zone - Larger */}
        <div
          {...getRootProps()}
          className={`relative border-2 border-dashed rounded-xl p-16 text-center cursor-pointer transition-all duration-300 ${
            isDragActive
              ? 'border-blue-500 bg-blue-50 scale-102'
              : 'border-gray-300 hover:border-gray-400 hover:bg-gray-50'
          } ${uploading || processing || uploadedFiles.length > 0 ? 'pointer-events-none opacity-50' : ''}`}
        >
          <input {...getInputProps()} />
          {uploading ? (
            <LoadingSpinner size="xl" text="Uploading files..." />
          ) : uploadedFiles.length > 0 ? (
            <div className="space-y-6">
              <div className="flex justify-center">
                <CheckCircle className="h-16 w-16 text-green-500" />
              </div>
              <div>
                <p className="text-2xl font-medium text-gray-900">
                  Files uploaded successfully
                </p>
                <p className="text-lg text-gray-500 mt-3">
                  Click "Clear All" below to upload different files
                </p>
              </div>
            </div>
          ) : (
            <div className="space-y-6">
              <div className="flex justify-center">
                <Upload className="h-16 w-16 text-gray-400" />
              </div>
              <div>
                <p className="text-2xl font-medium text-gray-900">
                  {isDragActive ? 'Drop files here' : 'Drag & drop your files'}
                </p>
                <p className="text-lg text-gray-500 mt-3">
                  or click to select files
                </p>
              </div>
              <div className="bg-blue-50 rounded-lg p-6 max-w-md mx-auto">
                <p className="text-base text-blue-800 font-medium">Requirements:</p>
                <ul className="text-base text-blue-700 mt-2 space-y-2">
                  <li>• Exactly one PDF presentation file</li>
                  <li>• Exactly one Excel source file</li>
                </ul>
              </div>
            </div>
          )}
        </div>

        {/* Enhanced Uploaded Files List */}
        {uploadedFiles.length > 0 && (
          <div className="bg-white shadow-lg rounded-xl overflow-hidden">
            <div className="bg-gradient-to-r from-blue-500 to-green-600 px-6 py-4">
              <h3 className="text-lg font-semibold text-white">
                Uploaded Files ({uploadedFiles.length})
              </h3>
            </div>
            
            <div className="p-6 space-y-4">
              {uploadedFiles.map((file) => (
                <div key={file.id} className="flex items-center justify-between p-4 bg-gray-50 rounded-lg border border-gray-200 hover:border-gray-300 transition-colors">
                  <div className="flex items-center space-x-4">
                    <div className={`p-3 rounded-lg ${
                      file.document_type === 'pdf' 
                        ? 'bg-red-100 text-red-600' 
                        : 'bg-green-100 text-green-600'
                    }`}>
                      {file.document_type === 'pdf' ? (
                        <FileText className="h-8 w-8" />
                      ) : (
                        <Table className="h-8 w-8" />
                      )}
                    </div>
                    <div>
                      <p className="text-sm font-medium text-gray-900">{file.filename}</p>
                      <p className="text-xs text-gray-500">
                        {(file.file_size / 1024 / 1024).toFixed(2)} MB • {file.document_type.toUpperCase()}
                      </p>
                      <div className="flex items-center mt-1">
                        <div className="w-2 h-2 bg-green-500 rounded-full mr-2"></div>
                        <span className="text-xs text-green-600 font-medium">Ready for direct validation</span>
                      </div>
                    </div>
                  </div>
                  
                  <div className="flex items-center space-x-2">
                    <CheckCircle className="h-6 w-6 text-green-500" />
                    <button
                      onClick={() => removeFile(file.id)}
                      className="text-gray-400 hover:text-red-500 transition-colors"
                      disabled={processing}
                    >
                      <X className="h-5 w-5" />
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Enhanced Action Buttons */}
        {uploadedFiles.length > 0 && (
          <div className="flex justify-center space-x-4">
            <button
              onClick={() => {
                setUploadedFiles([]);
                setSessionId(null);
              }}
              disabled={processing}
              className="px-8 py-3 border border-gray-300 rounded-lg text-sm font-medium text-gray-700 bg-white hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
            >
              Clear All
            </button>
            
            <button
              onClick={startComprehensiveProcessing}
              disabled={processing || uploading}
              className="px-8 py-3 border border-transparent rounded-lg shadow-sm text-sm font-medium text-white bg-gradient-to-r from-blue-600 to-green-600 hover:from-blue-700 hover:to-green-700 disabled:opacity-50 disabled:cursor-not-allowed flex items-center transition-all transform hover:scale-105"
            >
              Start Direct Validation Analysis
              <ArrowRight className="ml-2 h-4 w-4" />
            </button>
          </div>
        )}

        {/* Enhanced Instructions */}
        <div className="bg-gradient-to-r from-blue-50 to-green-50 border border-blue-200 rounded-xl p-8">
          <h3 className="text-xl font-semibold text-blue-900 mb-4 flex items-center justify-center">
            <Zap className="h-6 w-6 mr-2" />
            Direct Validation Approach
          </h3>
          <div className="grid md:grid-cols-2 gap-6">
            <div className="space-y-8 mt-4">
              <div className="flex items-start space-x-3">
                <div className="w-8 h-8 bg-blue-500 rounded-full flex items-center justify-center text-white text-sm font-bold">1</div>
                <div>
                  <h4 className="font-medium text-blue-900">Comprehensive Extraction</h4>
                  <p className="text-sm text-blue-800">Gemini 2.5 Pro extracts ALL values from PDF and Excel with complete business context</p>
                </div>
              </div>
              <div className="flex items-start space-x-3">
                <div className="w-8 h-8 bg-purple-500 rounded-full flex items-center justify-center text-white text-sm font-bold">2</div>
                <div>
                  <h4 className="font-medium text-purple-900">100% Coverage Audit</h4>
                  <p className="text-sm text-purple-800">AI audits ALL extracted values against each other for comprehensive validation</p>
                </div>
              </div>
            </div>
            <div className="space-y-4">
              <div className="bg-white rounded-lg p-4 border border-blue-200">
                <div className="flex items-center text-blue-600 text-sm font-medium mb-1">
                  <Brain className="h-4 w-4 mr-1" />
                  Powered by Orbicle Labs Pvt. Ltd.
                </div>
                <p className="text-xs text-blue-700">Advanced computer vision and business intelligence for enterprise-grade accuracy</p>
              </div>
              <div className="bg-white rounded-lg p-4 border border-purple-200">
                <div className="flex items-center text-purple-600 text-sm font-medium mb-1">
                  <Users className="h-4 w-4 mr-1" />
                  Human Auditor Workflow
                </div>
                <p className="text-xs text-purple-700">Follows real auditor process: validate extractions first, then compare everything</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </>
  );
};

export default EnhancedFileUpload;