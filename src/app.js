const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const morgan = require('morgan');
const path = require('path');
const { ROOT } = require('./utils/paths');

const uploadRoutes = require('./routes/upload.routes');
const insertosRoutes = require('./routes/insertos.routes');
const composeRoutes = require('./routes/compose.routes');
const composeOnceRoutes = require('./routes/compose-once.routes')
const inspectRoutes = require('./routes/inspect.routes');
const { ensureBaseDirs } = require('./services/storage.service');

const app = express();

app.use(helmet({
  hsts: false,                      // NO enviar Strict-Transport-Security
  crossOriginOpenerPolicy: false,
  crossOriginEmbedderPolicy: false,
  contentSecurityPolicy: {
    useDefaults: true,
    directives: {
      "default-src": ["'self'"],
      "script-src":  ["'self'", "'unsafe-inline'"],
      "style-src":   ["'self'", "'unsafe-inline'"],
      "img-src":     ["'self'", "data:"],
      "font-src":    ["'self'", "data:"],
      "connect-src": ["'self'"],
      "object-src":  ["'none'"],
      "base-uri":    ["'self'"],
      "frame-ancestors": ["'none'"],
      // MUY IMPORTANTE: NO pongas "upgrade-insecure-requests"
    }
  }
}));

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(morgan('dev'));

ensureBaseDirs();

// ⬇️ Servir frontend en la raíz "/"
app.use(express.static(path.join(ROOT, 'public'), {
  index: 'index.html',
  extensions: ['html']
}));
app.get('/', (_req, res) =>
  res.sendFile(path.join(ROOT, 'public', 'index.html'))
);

// Health
app.get('/health', (_req, res) => res.json({ ok: true }));

// ⬇️ API
app.use('/api/upload', uploadRoutes);
app.use('/api/insertos', insertosRoutes);
app.use('/api/compose', composeRoutes);
app.use('/api/compose-once', composeOnceRoutes); // nuevo (1 paso, en memoria)
app.use('/api/inspect', inspectRoutes);

module.exports = app;
