# Inlenersbeloning Excel Plugin

This Excel plugin loads remuneration data from the Inlenersbeloning API directly into Excel spreadsheets with automatic formatting.

## Features

- **Partner Selection**: Dropdown list of available partners from the API
- **Automatic Data Loading**: Fetches remuneration data for selected partners
- **Smart Error Handling**: Shows "No remuneration available" when appropriate
- **Excel Integration**: Automatically formats data as professional tables
- **Real-time Status**: Clear feedback on loading progress and errors

## API Integration

The plugin connects to the Inlenersbeloning API with the following endpoints:

1. **Partners**: `https://mijn.inlenersbeloning.com/api/v6/partner/rvkd2l`
   - Loads all available partners for the dropdown

2. **Partner Remuneration**: `https://mijn.inlenersbeloning.com/api/v6/remuneration/rvkd2l/{partnerId}`
   - Checks if remuneration data exists for a partner

3. **Detailed Remuneration**: `https://mijn.inlenersbeloning.com/api/v6/remuneration/rvkd2l/{partnerId}/{remunerationId}`
   - Loads detailed remuneration data for Excel insertion

Authentication is handled automatically using a Bearer token.

## Getting Started

### Prerequisites

- Microsoft Excel (Desktop or Online)
- Node.js and npm installed
- VS Code (recommended for development)

### Installation

1. Clone the repository:

   ```bash
   git clone <repository-url>
   cd ilb-excel-v1
   ```

2. Install dependencies:

   ```bash
   npm install
   ```

3. Build the project:

   ```bash
   npm run build:dev
   ```

4. Start the development server:

   ```bash
   npm run dev-server
   ```

5. Sideload the add-in in Excel:

   ```bash
   npm run start
   ```

## Usage

### Loading Data from APIs

1. **Custom API Endpoint**:
   - Enter your API URL in the "API URL" field
   - Optionally add custom headers in JSON format
   - Click "Load Custom API" to fetch and insert data

2. **Quick Examples**:
   - **Load Users**: Fetches user data from JSONPlaceholder API
   - **Load Posts**: Fetches blog posts data
   - **Exchange Rates**: Fetches current currency exchange rates

3. **Sample Data**: Click "Load Sample Data" to insert demonstration data

### API Service Features

The plugin includes a comprehensive API service (`src/services/apiService.ts`) with:

- **Request Methods**: GET, POST, PUT, DELETE
- **Timeout Handling**: Configurable request timeouts
- **Error Handling**: Detailed error responses
- **Authentication**: Support for various auth methods
- **Response Transformation**: Automatic JSON/text parsing

### Example API Calls

```typescript
import { apiService } from './services/apiService';

// Simple GET request
const response = await apiService.get('https://api.example.com/data');

// POST request with data
const response = await apiService.post('https://api.example.com/users', {
  name: 'John Doe',
  email: 'john@example.com'
});

// Request with custom headers
const response = await apiService.get('https://api.example.com/protected', {
  'Authorization': 'Bearer your-token-here'
});
```

### Data Transformation

The plugin automatically converts various data formats to Excel-compatible 2D arrays:

- **Array of Objects**: Converted to table with headers
- **Single Object**: Converted to key-value pairs
- **Nested Objects**: Flattened or JSON-stringified
- **Primitive Values**: Wrapped in array format

## File Structure

```text
src/
├── services/
│   ├── apiService.ts      # Core API service functionality
│   └── exampleApis.ts     # Example API implementations
├── taskpane/
│   ├── taskpane.ts        # Main taskpane logic
│   ├── taskpane.html      # UI layout
│   └── taskpane.css       # Styling
└── ...
```

## API Service Reference

### ApiService Class

```typescript
class ApiService {
  constructor(baseUrl?: string, defaultHeaders?: Record<string, string>)
  
  // HTTP Methods
  async get<T>(endpoint: string, headers?: Record<string, string>): Promise<ApiResponse<T>>
  async post<T>(endpoint: string, body?: any, headers?: Record<string, string>): Promise<ApiResponse<T>>
  async put<T>(endpoint: string, body?: any, headers?: Record<string, string>): Promise<ApiResponse<T>>
  async delete<T>(endpoint: string, headers?: Record<string, string>): Promise<ApiResponse<T>>
  
  // Configuration
  setBaseUrl(baseUrl: string): void
  setDefaultHeaders(headers: Record<string, string>): void
  setAuthToken(token: string, type?: string): void
}
```

### ApiResponse Interface

```typescript
interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  error?: string;
  status?: number;
}
```

## Example APIs Included

1. **JSONPlaceholder APIs**:
   - Users: `https://jsonplaceholder.typicode.com/users`
   - Posts: `https://jsonplaceholder.typicode.com/posts`

2. **Exchange Rates API**:
   - Current rates: `https://api.exchangerate-api.com/v4/latest/USD`

3. **Weather API** (requires API key):
   - OpenWeatherMap: `https://api.openweathermap.org/data/2.5/weather`

## Development

### Building the Project

```bash
# Development build
npm run build:dev

# Production build
npm run build

# Watch mode (rebuilds on changes)
npm run watch
```

### Testing

```bash
# Start dev server
npm run dev-server

# Debug in Excel Desktop
npm run start -- desktop --app excel

# Stop debugging
npm run stop
```

### Linting

```bash
# Check for issues
npm run lint

# Fix auto-fixable issues
npm run lint:fix
```

## Customization

### Adding New API Endpoints

1. Create a new function in `src/services/exampleApis.ts`:

```typescript
export async function fetchMyApiData(): Promise<any[][]> {
  try {
    const response = await apiService.get('https://my-api.com/data');
    
    if (!response.success) {
      throw new Error(response.error || 'Failed to fetch data');
    }

    // Transform data to Excel format
    return convertToExcelFormat(response.data);
  } catch (error) {
    console.error('Error fetching data:', error);
    return [['Error'], [error.message]];
  }
}
```

1. Add a button in `taskpane.html`:

```html
<div role="button" id="load-my-data" class="ms-Button ms-Button--secondary ms-font-s">
    <span class="ms-Button-label">Load My Data</span>
</div>
```

1. Wire up the event in `taskpane.ts`:

```typescript
document.getElementById("load-my-data").onclick = loadMyData;

export async function loadMyData() {
  updateStatus("Loading my data...", "loading");
  try {
    await Excel.run(async (context) => {
      const data = await fetchMyApiData();
      await insertDataIntoExcel(context, data);
      updateStatus("My data loaded successfully", "success");
    });
  } catch (error) {
    updateStatus(`Error: ${error.message}`, "error");
  }
}
```

## Troubleshooting

### Common Issues

1. **CORS Errors**: Many APIs don't allow browser requests. Consider using a proxy server or server-side integration.

2. **Authentication**: For APIs requiring authentication, add headers in the Headers field:

   ```json
   {"Authorization": "Bearer your-token-here"}
   ```

3. **Large Datasets**: Excel has limits on cell count. Consider pagination or data filtering for large datasets.

4. **Network Timeouts**: Increase timeout in API service if needed:

   ```typescript
   const response = await apiService.request(url, { timeout: 60000 });
   ```

### Error Messages

- "Please enter an API URL": The URL field is empty
- "Network error occurred": Connection issues or invalid URL
- "HTTP 4XX/5XX": Server returned an error status
- "Request timeout": API took too long to respond

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## Support

For issues and questions:

- Check the GitHub Issues page
- Review the Excel Add-ins documentation
- Consult the Office JavaScript API reference
