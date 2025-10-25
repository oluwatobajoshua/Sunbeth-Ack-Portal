# Sunbeth ACK Portal

[![React](https://img.shields.io/badge/React-18.2.0-blue?logo=react)](https://reactjs.org/)
[![TypeScript](https://img.shields.io/badge/TypeScript-4.9.5-blue?logo=typescript)](https://www.typescriptlang.org/)
[![Azure](https://img.shields.io/badge/Azure-MSAL-blue?logo=microsoftazure)](https://azure.microsoft.com/)
[![License](https://img.shields.io/badge/license-Private-red)](./LICENSE)

A comprehensive web portal built with React and TypeScript, featuring Azure Active Directory integration, role-based access control (RBAC), document management, and business process automation.

## 🚀 Features

### Authentication & Security
- **Azure AD Integration**: Single Sign-On (SSO) with Microsoft Azure Active Directory
- **Multi-Factor Authentication (MFA)**: Enhanced security with TOTP support
- **Role-Based Access Control (RBAC)**: Fine-grained permissions and access control
- **External User Management**: Support for guest users and external authentication
- **Token Management**: Secure token handling and refresh mechanisms

### Document Management
- **Document Viewer**: Built-in support for PDF and DOCX files
- **Bulk Upload**: Mass import of documents and data
- **Document Processing**: Automated document parsing and processing
- **SharePoint Integration**: Seamless integration with Microsoft SharePoint

### Business Process Management
- **Batch Processing**: Automated batch creation and management
- **Workflow Management**: Streamlined business process automation
- **Audit Logging**: Comprehensive activity tracking and audit trails
- **Analytics Dashboard**: Real-time insights and reporting

### User Experience
- **Responsive Design**: Mobile-friendly interface
- **Toast Notifications**: User-friendly feedback system
- **Modal Dialogs**: Intuitive user interactions
- **Loading States**: Enhanced UX with loading indicators

## 🏗️ Architecture

### Frontend Stack
- **React 18.2**: Modern React with concurrent features
- **TypeScript 4.9**: Type-safe development
- **React Router 6**: Client-side routing
- **Azure MSAL**: Microsoft Authentication Library
- **Microsoft Graph Toolkit**: Office 365 integration

### Backend Services
- **Express.js**: RESTful API server
- **SQLite**: Local database for development
- **Node.js**: Server-side JavaScript runtime

### Key Libraries
- **Document Processing**: `docx-preview`, `react-pdf`, `pdfjs-dist`
- **Data Handling**: `xlsx`, `papaparse` for Excel/CSV processing
- **UI Components**: Custom React components with TypeScript
- **Testing**: Jest, React Testing Library

## 📋 Prerequisites

Before running this project, ensure you have:

- **Node.js** (v16 or higher)
- **npm** or **yarn**
- **Azure AD** application registration
- **SharePoint** access (optional)

## 🛠️ Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/oluwatobajoshua/Sunbeth-Ack-Portal.git
   cd Sunbeth-Ack-Portal
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Install server dependencies**
   ```bash
   cd server
   npm install
   cd ..
   ```

4. **Configure environment variables**
   Create a `.env` file in the root directory:
   ```env
   # Azure AD Configuration
   REACT_APP_AZURE_CLIENT_ID=your_client_id
   REACT_APP_AZURE_TENANT_ID=your_tenant_id
   REACT_APP_AZURE_REDIRECT_URI=http://localhost:3000

   # API Configuration (optional for dev; proxy targets http://localhost:4000)
   REACT_APP_API_BASE_URL=http://localhost:4000

   # Feature Flags
   REACT_APP_ENABLE_EXTERNAL_AUTH=true
   REACT_APP_ENABLE_MFA=true
   ```

## 🚀 Getting Started

### Development Mode

1. **Start the frontend development server**
   ```bash
   npm start
   ```
   The application will open at `http://localhost:3000`

2. **Start the backend API server** (in a separate terminal)
   ```bash
   npm run api
   ```
   The API will be available at `http://localhost:4000`
   
   Development requests to `/api/*` from the React app are proxied to `http://127.0.0.1:4000` by `src/setupProxy.js`.

### Production Build

1. **Build the application**
   ```bash
   npm run build
   ```

2. **Preview the production build**
   ```bash
   npm run preview
   ```
   The static build is served at `http://localhost:5000`. If your backend is required, run `npm run api` (on port 4000) separately.

## 📁 Project Structure

```
src/
├── components/           # React components
│   ├── About.tsx
│   ├── AdminPanel.tsx
│   ├── AnalyticsDashboard.tsx
│   ├── BatchCreationDebug.tsx
│   ├── BusinessesBulkUpload.tsx
│   ├── Dashboard.tsx
│   ├── DocumentReader.tsx
│   ├── ExternalLogin.tsx
│   ├── LoginGateway.tsx
│   └── viewers/         # Document viewer components
├── context/             # React Context providers
│   ├── AuthContext.tsx
│   ├── FeatureFlagsContext.tsx
│   └── RBACContext.tsx
├── services/            # API and service layers
│   ├── authInteractive.ts
│   ├── dbService.ts
│   ├── graphService.ts
│   ├── sharepointService.ts
│   └── msalConfig.ts
├── utils/               # Utility functions
│   ├── alerts.ts
│   ├── batchLogger.ts
│   ├── excelExport.ts
│   └── importTemplates.ts
├── types/               # TypeScript type definitions
│   └── models.ts
└── diagnostics/         # Error handling and logging
    ├── ErrorBoundary.tsx
    ├── logger.ts
    └── health.ts
```

## 🔧 Available Scripts

- `npm start` - Start development server
- `npm run build` - Build for production
- `npm test` - Run test suite
- `npm run preview` - Preview production build
- `npm run api` - Start backend API server
- `npm run docs` - Generate TypeScript documentation
- `npm run fix:sourcemaps` - Fix sourcemap issues

## 🧪 Testing

Run the test suite:
```bash
npm test
```

The project includes:
- Unit tests with Jest
- Component tests with React Testing Library
- Navigation and UX tests

## 🔐 Authentication Setup

### Azure AD Configuration

1. **Register your application** in Azure AD
2. **Configure redirect URIs** for your environment
3. **Set API permissions** for Microsoft Graph
4. **Update MSAL configuration** in `src/services/msalConfig.ts`

### RBAC Configuration

The application supports role-based access control with the following roles:
- **Admin**: Full system access
- **Manager**: Business process management
- **User**: Standard user access
- **Guest**: Limited external access

## 📊 Features Overview

### Dashboard
- Real-time analytics and metrics
- Batch processing status
- User activity summaries
- System health monitoring

### Batch Management
- Create and manage processing batches
- Monitor batch progress and status
- Debug batch creation issues
- Export batch results

### User Management
- Internal user authentication via Azure AD
- External user invitation and management
- Role assignment and permissions
- Bulk user import capabilities

### Document Processing
- Upload and process various document types
- PDF and DOCX document viewers
- Automated document parsing
- Integration with SharePoint document libraries

## 🚢 Deployment

### Recommended Platforms

The application can also be deployed to:
- **Azure Static Web Apps**
- **Netlify**
- **Vercel**
- **Traditional web servers**

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Guidelines

- Follow TypeScript best practices
- Write tests for new features
- Use consistent code formatting
- Update documentation for significant changes
- Follow the existing project structure

## 📝 License

This project is private and proprietary. All rights reserved.

## 🐛 Troubleshooting

### Common Issues

**Authentication Issues**
- Verify Azure AD configuration
- Check redirect URIs match your environment
- Ensure proper API permissions are granted

**Build Issues**
- Clear node_modules and reinstall: `rm -rf node_modules package-lock.json && npm install`
- Check TypeScript version compatibility
- Run `npm run fix:sourcemaps` if encountering sourcemap issues

**API Connection Issues**
- Verify backend server is running
- Check CORS configuration
- Confirm API endpoint URLs

### Getting Help

- Check the [Issues](https://github.com/oluwatobajoshua/Sunbeth-Ack-Portal/issues) page
- Review the TypeScript documentation: `npm run docs`
- Contact the development team

## � Project Status

### ✅ Completed Features

**Authentication & Security**
- ✅ Azure AD integration with MSAL
- ✅ Multi-Factor Authentication (MFA) with TOTP
- ✅ External user authentication system
- ✅ Role-Based Access Control (RBAC)
- ✅ Token management and refresh
- ✅ Login gateway and unified authentication

**Core Application**
- ✅ React 18 with TypeScript setup
- ✅ Responsive layout and navigation
- ✅ Context providers for auth, RBAC, and feature flags
- ✅ Error boundaries and diagnostics
- ✅ Toast notification system
- ✅ Modal dialog system
- ✅ Loading states and skeleton screens

**Document Management**
- ✅ PDF viewer integration (`react-pdf`, `pdfjs-dist`)
- ✅ DOCX viewer integration (`docx-preview`)
- ✅ Document reader component
- ✅ File upload capabilities

**Business Process Management**
- ✅ Dashboard with batch overview
- ✅ Batch creation and management
- ✅ Progress tracking system
- ✅ Batch completion workflow
- ✅ Debug console for batch creation
- ✅ Comprehensive batch logging system

**Data Management**
- ✅ SQLite database integration
- ✅ Excel/CSV import functionality (`xlsx`, `papaparse`)
- ✅ Business entity CRUD operations
- ✅ External user management
- ✅ Bulk upload systems

**Backend Services**
- ✅ Express.js API server
- ✅ SQLite database operations
- ✅ User authentication endpoints
- ✅ Business data API
- ✅ File upload handling

### 🚧 Partially Implemented Features

**SharePoint Integration**
- ⚠️ Configuration structure in place
- ❌ CRUD operations not implemented
- ❌ Document library integration pending
- ❌ List management functionality missing

**Analytics & Reporting**
- ✅ Basic dashboard analytics
- ✅ Progress tracking metrics
- ⚠️ Advanced reporting features limited
- ❌ Export functionality incomplete

**Testing**
- ✅ Test framework setup (Jest, React Testing Library)
- ✅ Basic navigation tests
- ⚠️ Component test coverage incomplete
- ❌ End-to-end tests missing

### ❌ Pending Implementation

**SharePoint Lists Backend**
```typescript
// From dbService.ts - These operations throw "not implemented" errors:
- Business creation for SharePoint Lists
- Business updates for SharePoint Lists  
- Business deletion for SharePoint Lists
```

**Token Management Enhancement**
```typescript
// From ExternalLogin.tsx:
// TODO: Store token/session as needed
```

**Advanced Features**
- Workflow automation system
- Advanced document processing
- Real-time notifications
- Audit trail enhancements
- Performance monitoring
- Mobile responsiveness optimization

**Infrastructure**
- Production deployment pipeline
- Environment-specific configurations
- Monitoring and logging infrastructure

### 🎯 Current Development Priorities

- Optimize SQLite + Express data flows
- Extend SharePoint integration
- Improve test coverage and E2E tests
- Harden RBAC and auditing

## 📈 Future Roadmap

### Phase 1 (Immediate - Next 1-2 weeks) 🔥
- [ ] Strengthen API error handling and validation
- [ ] SharePoint Lists CRUD integration
- [ ] Expand document processing pipeline
- [ ] Improve CI checks and deployment scripts

### Phase 2 (Short-term - Next 1-3 months)
- [ ] Enhanced analytics and reporting
- [ ] Advanced document processing
- [ ] Workflow automation improvements
- [ ] Mobile responsiveness enhancements
- [ ] Performance optimizations

### Phase 3 (Medium-term - Next 3-6 months)
- [ ] Mobile application
- [ ] Additional authentication providers
- [ ] Real-time collaboration features
- [ ] Advanced audit and compliance features
- [ ] API rate limiting and security enhancements

### Phase 4 (Long-term - 6+ months)
- [ ] Machine learning integration for document processing
- [ ] Advanced workflow designer
- [ ] Multi-tenant architecture
- [ ] Internationalization and localization
- [ ] Enterprise integrations

---

**Built with ❤️ by the Sunbeth Development Team**