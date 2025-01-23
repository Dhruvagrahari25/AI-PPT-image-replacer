import React from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import './index.css';
import App from './App';
import Profile from './pages/Profile';
import HomePage from './pages/HomePage';


ReactDOM.createRoot(document.getElementById('root')).render(
  <Router>
    <Routes>
      <Route path="/" element={<HomePage />} />
      <Route path="/forum" element={<App />} />
      <Route path="/profile" element={<Profile />} />
    </Routes>
  </Router>
);
