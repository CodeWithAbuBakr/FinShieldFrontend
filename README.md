# FinShield Outlook Add-in

## Overview
FinShield is a powerful Outlook add-in designed to enhance email security, provide financial insights, and integrate seamlessly into your workflow. Built with a robust frontend using React and a secure backend powered by Node.js, FinShield ensures a smooth and efficient experience.

## Features
- **Secure Email Tracking & Verification** – Ensures authenticity and integrity of emails.
- **Financial Insights Integration** – Provides valuable insights directly within Outlook.
- **Microsoft Single Sign-On (SSO)** – Seamless authentication for users.
- **Email Categorization (UI/Background Processing)** – Classifies emails based on risk levels.
- **Pre-send Validation** – Prevents risky emails from being sent.
- **Pinable Taskpane** – Persistent access via the Outlook manifest.
- **Secure API Communication** – Ensures encrypted and safe data transmission.
- **Modern & Responsive UI** – Built with Fluent UI components for a seamless user experience.

---

## Frontend (Outlook Add-in)
### Tech Stack
- React.js
- TypeScript
- Fluent UI
- Office.js
- Microsoft Graph API
- MSAL.js for authentication

### Installation & Setup
#### Prerequisites
- Node.js installed
- Office 365 Outlook
- Outlook desktop/web with add-ins enabled

#### Steps
1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/finshield-outlook-frontend.git
   cd finshield-outlook-frontend
   ```
2. **Install dependencies:**
   ```bash
   npm install
   ```
3. **Set up environment variables** in a `.env` file:
   ```env
   REACT_APP_CLIENT_ID=your_microsoft_app_client_id
   REACT_APP_BACKEND_URL=http://localhost:4000
   ```
4. **Start the development server:**
   ```bash
   npm start
   ```
5. **Sideload the add-in in Outlook:**
   - Open Outlook Web or Desktop.
   - Navigate to **Settings > Manage Add-ins > Upload Manifest**.
   - Select the `manifest.xml` file from the project directory.

---

## Backend (API Server)
### Tech Stack
- Node.js
- Express.js
- MSAL for authentication
- Microsoft Graph API
- JWT for token handling

### Installation & Setup
#### Prerequisites
- Node.js installed

#### Steps
1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/finshield-backend.git
   cd finshield-backend
   ```
2. **Install dependencies:**
   ```bash
   npm install
   ```
3. **Set up environment variables** in a `.env` file:
   ```env
   PORT=4000
   JWT_SECRET=your_secret_key
   AZURE_CLIENT_ID=your_microsoft_client_id
   AZURE_TENANT_ID=your_tenant_id
   ```
4. **Start the backend server:**
   ```bash
   npm start
   ```

---

## Deployment
### Frontend Deployment
- Deploy on **GitHub Pages, Vercel, or Azure Static Web Apps**.
- Ensure the backend URL is correctly set in `.env.production`.

### Backend Deployment
- Deploy on **Heroku, Vercel, or Azure App Service**.
- Use a cloud database like **MongoDB Atlas** for persistent storage.

---

## API Endpoints
| Method | Endpoint                          | Description                     |
|--------|-----------------------------------|---------------------------------|
| POST   | `/api/login/auth/microsoft`       | User authentication             |
| POST   | `/api/login/auth/refresh`         | Refresh authentication token    |

---

## License
This project is licensed under the **MIT License**.

## Contact
For support or inquiries, contact **SAAD AFZAL**.

