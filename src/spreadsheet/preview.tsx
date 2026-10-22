import '../react/polyfills/buffer';
import React from 'react';
import {createRoot} from 'react-dom/client';
import Excel from '../react/view/excel/Excel';
import './preview.css';

// Keep the complete upstream spreadsheet application. Compatibility and native
// calculation are added beneath it, rather than replacing its toolbar/editor.
createRoot(document.getElementById('root')!).render(<Excel/>);
