import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { registerLicense } from '@syncfusion/ej2-base';
import '@syncfusion/ej2-base/styles/material3.css';
import '@syncfusion/ej2-buttons/styles/material3.css';
import '@syncfusion/ej2-dropdowns/styles/material3.css';
import '@syncfusion/ej2-inputs/styles/material3.css';
import '@syncfusion/ej2-lists/styles/material3.css';
import '@syncfusion/ej2-navigations/styles/material3.css';
import '@syncfusion/ej2-popups/styles/material3.css';
import '@syncfusion/ej2-splitbuttons/styles/material3.css';
import '@syncfusion/ej2-pdfviewer/styles/material3.css';
import '@syncfusion/ej2-react-documenteditor/styles/material3.css';
import '@syncfusion/ej2-documenteditor/styles/material3.css';
import './index.css';
import App from './App';

const licenseKey = import.meta.env.VITE_SYNCFUSION_LICENSE_KEY as string | undefined;
if (licenseKey) {
  registerLicense(licenseKey);
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
