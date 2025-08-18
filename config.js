// Configuration for API URLs
const config = {
    // Development (local)
    development: {
        apiBaseUrl: 'http://localhost:5000'
    },
    // Production (deployed)
    production: {
        apiBaseUrl: 'https://kaam-asaasan-api.herokuapp.com' // Will be updated after deployment
    }
};

// Get current environment
const isProduction = window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1';
const currentConfig = isProduction ? config.production : config.development;

// Export for use in other files
window.API_CONFIG = currentConfig;
