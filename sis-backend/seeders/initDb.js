/**
 * Create DB if needed and apply schema.
 * Set RESET_DB=1 to drop public schema and recreate (fixes old mismatched tables).
 *
 *   RESET_DB=1 node seeders/initDb.js
 *   npm run db:init
 */
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { Client } = require('pg');

function baseConfig(database) {
  return {
    host: process.env.DB_HOST || process.env.PGHOST || 'localhost',
    port: Number(process.env.DB_PORT || process.env.PGPORT || 5432),
    user: process.env.DB_USER || process.env.PGUSER || 'postgres',
    password: process.env.DB_PASSWORD || process.env.PGPASSWORD || 'postgres',
    database,
  };
}

async function main() {
  const dbName = process.env.DB_NAME || process.env.PGDATABASE || 'sis_db';
  const reset = process.env.RESET_DB === '1' || process.argv.includes('--reset');

  const admin = new Client(baseConfig('postgres'));
  await admin.connect();
  const exists = await admin.query('SELECT 1 FROM pg_database WHERE datname = $1', [dbName]);
  if (!exists.rowCount) {
    await admin.query(`CREATE DATABASE ${dbName}`);
    console.log('Created database', dbName);
  } else {
    console.log('Database exists', dbName);
  }
  await admin.end();

  const client = new Client(baseConfig(dbName));
  await client.connect();

  if (reset) {
    console.log('RESET_DB: dropping public schema...');
    await client.query('DROP SCHEMA IF EXISTS public CASCADE');
    await client.query('CREATE SCHEMA public');
    await client.query('GRANT ALL ON SCHEMA public TO public');
    console.log('Public schema recreated');
  }

  const schemaPath = path.join(__dirname, '..', 'sql', 'schema.sql');
  const schema = fs.readFileSync(schemaPath, 'utf8');
  await client.query(schema);
  console.log('Schema applied to', dbName);
  await client.end();
  console.log('Done. Next: npm run db:seed');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
