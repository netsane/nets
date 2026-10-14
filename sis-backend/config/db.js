const { Pool } = require('pg');
require('dotenv').config();

/**
 * Supports both naming styles:
 *   PGHOST / PGUSER / PGPASSWORD / PGDATABASE / PGPORT  (unity template)
 *   DB_HOST / DB_USER / DB_PASSWORD / DB_NAME / DB_PORT / DATABASE_URL (your sis-backend)
 */
const connectionString = process.env.DATABASE_URL || process.env.DB_URL || null;

const pool = connectionString
  ? new Pool({ connectionString })
  : new Pool({
      host: process.env.DB_HOST || process.env.PGHOST || 'localhost',
      port: Number(process.env.DB_PORT || process.env.PGPORT || 5432),
      user: process.env.DB_USER || process.env.PGUSER || 'postgres',
      password: process.env.DB_PASSWORD || process.env.PGPASSWORD || 'postgres',
      database: process.env.DB_NAME || process.env.PGDATABASE || 'sis_db',
      max: 20,
      idleTimeoutMillis: 30000,
    });

pool.on('error', (err) => {
  console.error('[pg] idle client error', err.message);
});

async function query(text, params) {
  const start = Date.now();
  const res = await pool.query(text, params);
  if (process.env.NODE_ENV === 'development') {
    const ms = Date.now() - start;
    if (ms > 200) console.log('[pg]', ms, 'ms', text.slice(0, 80));
  }
  return res;
}

module.exports = { pool, query };
