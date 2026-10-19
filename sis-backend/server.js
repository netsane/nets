require('dotenv').config();
const express = require('express');
const cors = require('cors');
const apiRoutes = require('./routes');

const app = express();
const PORT = Number(process.env.PORT || 3001);

// Allow Vite on localhost and LAN IPs (phone/tablet testing)
const defaultOrigins = [
  'http://localhost:5174',
  'http://127.0.0.1:5174',
  'http://192.168.1.23:5174',
];
const envOrigins = (process.env.CORS_ORIGIN || '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);
const allowList = [...new Set([...defaultOrigins, ...envOrigins])];

app.use(
  cors({
    origin(origin, cb) {
      // allow same-origin / curl / server-to-server (no Origin header)
      if (!origin) return cb(null, true);
      if (allowList.includes(origin)) return cb(null, true);
      // any private LAN Vite host on port 5173
      if (/^http:\/\/(192\.168\.\d+\.\d+|10\.\d+\.\d+\.\d+|172\.(1[6-9]|2\d|3[0-1])\.\d+\.\d+):5173$/.test(origin)) {
        return cb(null, true);
      }
      return cb(null, false);
    },
    credentials: true,
  })
);
app.use(express.json({ limit: '10mb' }));

// All Unity SIS APIs under /api
app.use('/api', apiRoutes);

app.get('/', (_req, res) => {
  res.json({
    name: 'Unity SIS Backend',
    api: '/api',
    health: '/api/health',
  });
});

app.use((err, _req, res, _next) => {
  console.error(err);
  res.status(500).json({ success: false, message: 'Server error' });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`SIS API running → http://localhost:${PORT}/api`);
  console.log(`Health          → http://localhost:${PORT}/api/health`);
});
