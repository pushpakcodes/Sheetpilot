import axios from 'axios';

const API_URL = 'http://localhost:5000/api';

const api = axios.create({
  baseURL: API_URL,
});

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('token');
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

export const login = (email, password) => api.post('/auth/login', { email, password });
export const register = (username, email, password) => api.post('/auth/register', { username, email, password });
export const uploadFile = (formData) => api.post('/excel/upload', formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
});
export const processCommand = (command, filePath) => api.post('/excel/process', { command, filePath });
export const getFiles = () => api.get('/excel/files');

export default api;
