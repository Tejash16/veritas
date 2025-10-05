import axios from 'axios';

const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:8008/api';

class EnhancedApiService {
  constructor() {
    this.client = axios.create({
      baseURL: API_BASE_URL,
      timeout: 1000000,
    });

    // Request interceptor
    this.client.interceptors.request.use(
      (config) => {
        const token = localStorage.getItem('veritas_token');
        if (token) {
          config.headers.Authorization = `Bearer ${token}`;
        }

        // console.log(`[API] ${config.method?.toUpperCase()} ${config.url}`, {
        //   data: config.data,
        //   params: config.params
        // });

        return config;
      },
      (error) => {
        // console.error('[API] Request error:', error);
        return Promise.reject(error);
      }
    );

    // Response interceptor
    this.client.interceptors.response.use(
      (response) => {
        // console.log(`[API] Response:`, response.data);
        return response.data;
      },
      (error) => {
        // console.error('[API] Response error:', error);

        if (error.response?.status === 401) {
          // localStorage.removeItem('veritas_token');
          // window.location.href = '/';
        }

        const message = error.response?.data?.detail || error.message || 'An error occurred';
        throw new Error(message);
      }
    );
  }

  // Authentication
  async login(credentials) {
    const response = await this.client.post('/auth/login', credentials);
    return response;
  }

  async validateToken() {
    const response = await this.client.get('/auth/validate');
    return response;
  }

  // Enhanced document operations
  async uploadDocuments(formData) {
    const response = await this.client.post('/upload/documents', formData, {
      headers: {
        'Content-Type': 'multipart/form-data',
      },
    });
    return response;
  }

  // Processing - Direct Validation Approach
  async processDocuments(sessionId) {
    const response = await this.client.post(`/process/comprehensive/${sessionId}`);
    return response;
  }

  async processDocumentsComprehensive(sessionId) {
    const response = await this.client.post(`/process/comprehensive/${sessionId}`);
    return response;
  }

  // Direct Validation Operations
  async getValidationData(sessionId) {
    const response = await this.client.get(`/validation/data/${sessionId}`);
    return response;
  }

  async updatePdfValue(sessionId, valueData) {
    const response = await this.client.post(`/validation/update-pdf-value/${sessionId}`, valueData);
    return response;
  }

  async updateExcelValue(sessionId, valueData) {
    const response = await this.client.post(`/validation/update-excel-value/${sessionId}`, valueData);
    return response;
  }

  async startDirectAudit(sessionId) {
    const response = await this.client.post(`/validation/start-direct-audit/${sessionId}`);
    return response;
  }

  async getValidationStatus(sessionId) {
    const response = await this.client.get(`/validation/status/${sessionId}`);
    return response;
  }

  // Audit Results
  async getAuditResults(auditSessionId) {
    const response = await this.client.get(`/audit/results/${auditSessionId}`);
    return response;
  }

  // Statistics
  async getStats() {
    const response = await this.client.get('/documents/stats');
    return response;
  }

  // Health check
  async healthCheck() {
    const response = await this.client.get('/health');
    return response;
  }

  async getExcelMeta(fileId) {
    const response = await this.client.get(`/excel/meta`, {
      params: { file_id: fileId }
    });
    return response;
  }

  async getExcelPage({ fileId, sheet, r0, r1, c0, c1 }) {
    const response = await this.client.get(`/excel/page`, {
      params: {
        file_id: fileId,
        sheet,
        r0,
        r1,
        c0,
        c1
      }
    });
    return response;
  }

  async getExcelSpotlight({ fileId, sheet, cell }) {
    const response = await this.client.get(`/excel/spotlight`, {
      params: {
        file_id: fileId,
        sheet,
        cell
      }
    });
    return response;
  }
}

// Export both names for backward compatibility
const apiService = new EnhancedApiService();
const enhancedApiService = apiService;

export { apiService, enhancedApiService };
export default apiService;

export async function getExcelMeta(fileId) {
  return apiService.getExcelMeta(fileId);
}

export async function getExcelPage(params) {
  return apiService.getExcelPage(params);
}

export async function getExcelSpotlight(params) {
  return apiService.getExcelSpotlight(params);
}
