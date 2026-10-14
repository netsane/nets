const express = require('express');
const auth = require('../controllers/authController');
const crud = require('../controllers/crudController');
const integration = require('../controllers/integrationController');
const { authRequired, optionalAuth, requireModule } = require('../middleware/auth');

const router = express.Router();

/** Browser-friendly index — fixes "Cannot GET /api/" */
router.get('/', (_req, res) => {
  res.json({
    ok: true,
    service: 'Unity SIS Backend',
    message: 'API is running. Use the paths below (most need Authorization: Bearer <token>).',
    endpoints: {
      health: 'GET /api/health',
      login: 'POST /api/auth/login',
      me: 'GET /api/auth/me',
      bootstrap: 'GET /api/bootstrap',
      collections: {
        list: 'GET /api/collections/:name',
        one: 'GET /api/collections/:name/:id',
        create: 'POST /api/collections/:name',
        update: 'PUT /api/collections/:name/:id',
        remove: 'DELETE /api/collections/:name/:id',
        examples: [
          'HR_EMPLOYEE_CRUD',
          'FIN_COA_CRUD',
          'PAY_PROCESS_CRUD',
          'MAT_GRV_CRUD',
          'transactions',
        ],
      },
      integration: {
        syncHrPayMaster: 'POST /api/integration/hr/sync-pay-master',
        materialPost: 'POST /api/integration/material/post',
        payrollPost: 'POST /api/integration/payroll/post',
      },
    },
    tryInBrowser: ['http://localhost:3001/api/health', 'http://localhost:3001/', 'http://192.168.1.19:3001/'],
  });
});

router.get('/health', (_req, res) => res.json({ ok: true, service: 'sis-backend' }));

router.post('/auth/login', auth.login);
router.get('/auth/me', authRequired, auth.me);

router.get('/bootstrap', optionalAuth, crud.bootstrap);

router.get('/collections/:collection', optionalAuth, crud.list);
router.get('/collections/:collection/:id', authRequired, crud.getOne);
router.post('/collections/:collection', optionalAuth, crud.create);
router.put('/collections/:collection/:id', optionalAuth, crud.update);
router.delete('/collections/:collection/:id', optionalAuth, crud.remove);

router.post('/integration/material/post', authRequired, requireModule('material', 'finance'), integration.postMaterial);
router.post('/integration/hr/sync-pay-master', authRequired, requireModule('hr', 'finance'), integration.syncHrPay);
router.post('/integration/payroll/post', authRequired, requireModule('finance'), integration.postPayroll);

module.exports = router;
