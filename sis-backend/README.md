# SIS Backend — Unity Finance · HR · Material

## 1. Install & database

```bash
cd sis-backend
npm install

# ensure PostgreSQL is running and password matches .env (sisdb)
npm run db:init
npm run db:seed
npm run dev
```

API: **http://localhost:3001/api**

## 2. Your .env (already set)

```
PORT=3001
DB_HOST=localhost
DB_USER=postgres
DB_PASSWORD=sisdb
DB_NAME=sis_db
DB_PORT=5432
DATABASE_URL=postgresql://postgres:sisdb@localhost:5432/sis_db
JWT_SECRET=...
CORS_ORIGIN=http://localhost:5173
```

## 3. Default users (password: `password`)

- admin
- hr_manager
- finance_officer
- store_manager

## 4. Frontend

`sis-frontend/.env`:
```
VITE_API_BASE_URL=http://localhost:3001/api
```

Copy `frontend-connect/api.js` → `sis-frontend/src/unity/api.js` (or `src/api.js`).

## 5. Endpoints

| Method | Path |
|--------|------|
| POST | /api/auth/login |
| GET | /api/bootstrap |
| GET/POST | /api/collections/:name |
| PUT/DELETE | /api/collections/:name/:id |
| POST | /api/integration/hr/sync-pay-master |
| POST | /api/integration/material/post |
| POST | /api/integration/payroll/post |
