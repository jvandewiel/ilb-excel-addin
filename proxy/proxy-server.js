const express = require('express');
const cors = require('cors');
const axios = require('axios');

const app = express();
const PORT = 3001;

// API Configuration
const API_BASE_URL = 'https://mijn.inlenersbeloning.com/api/v6';
const API_TOKEN = 'org_fe7686f7bca2c1f1c79400ae6fc6ffc1673889b8aceb7f51f8a9d0f7d62e1e07';

// Enable CORS for all routes
app.use(cors());
app.use(express.json());

// Create axios instance with default headers
const apiClient = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Authorization': `Bearer ${API_TOKEN}`,
    'Content-Type': 'application/json'
  }
});

// Proxy routes
app.get('/api/v6/*', async (req, res) => {
  try {
    const path = req.path.replace('/api/v6', '');
    console.log(`Proxying GET request to: ${API_BASE_URL}${path}`);
    
    const response = await apiClient.get(path);
    
    console.log(`Response status: ${response.status}`);
    console.log(`Response data:`, response.data);
    
    res.json(response.data);
  } catch (error) {
    console.error('Proxy error:', error.message);
    
    if (error.response) {
      // API returned an error response
      res.status(error.response.status).json({
        error: error.response.data || error.message,
        status: error.response.status
      });
    } else {
      // Network or other error
      res.status(500).json({
        error: error.message,
        status: 500
      });
    }
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'Proxy server is running' });
});

app.listen(PORT, () => {
  console.log(`🚀 Proxy server running on http://localhost:${PORT}`);
  console.log(`📡 Proxying requests to: ${API_BASE_URL}`);
  console.log(`🔑 Using token: ${API_TOKEN.substring(0, 20)}...`);
});