import 'dotenv/config';
import express from 'express';
import apiHandler from './api.js';

const app = express();
app.use(express.json({ limit: '20mb' }));
app.use('/api', (req, res) => {
  req.url = req.originalUrl;
  apiHandler(req, res);
});
app.use((error, _req, res, _next) => {
  console.error('Local API error:', error);
  res.status(error.status || 500).json({ error: error.message || 'Internal server error.' });
});

const port = Number(process.env.PORT || 3001);
app.listen(port, () => {
  console.log(`Islam Marble API listening on http://localhost:${port}`);
});
