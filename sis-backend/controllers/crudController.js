const crud = require('../services/crudService');

async function list(req, res) {
  try {
    const data = await crud.listCollection(req.params.collection);
    res.json({ success: true, data });
  } catch (e) {
    console.error(e);
    res.status(500).json({ success: false, message: e.message });
  }
}

async function getOne(req, res) {
  try {
    const row = await crud.getById(req.params.collection, req.params.id);
    if (!row) return res.status(404).json({ success: false, message: 'Not found' });
    res.json({ success: true, data: row });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message });
  }
}

async function create(req, res) {
  try {
    const row = await crud.create(req.params.collection, req.body || {});
    res.status(201).json({ success: true, data: row });
  } catch (e) {
    console.error(e);
    res.status(500).json({ success: false, message: e.message });
  }
}

async function update(req, res) {
  try {
    const row = await crud.update(req.params.collection, req.params.id, req.body || {});
    if (!row) return res.status(404).json({ success: false, message: 'Not found' });
    res.json({ success: true, data: row });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message });
  }
}

async function remove(req, res) {
  try {
    const ok = await crud.remove(req.params.collection, req.params.id);
    if (!ok) return res.status(404).json({ success: false, message: 'Not found' });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message });
  }
}

/** Bulk dump for frontend bootstrap (all mapped collections) */
async function bootstrap(req, res) {
  try {
    const { COLLECTION_MAP } = require('../config/collections');
    const names = Object.keys(COLLECTION_MAP);
    const allData = {};
    for (const name of names) {
      allData[name] = await crud.listCollection(name);
    }
    // alias
    allData.chartsOfAccounts = allData.FIN_COA_CRUD || [];
    res.json({ success: true, allData });
  } catch (e) {
    console.error(e);
    res.status(500).json({ success: false, message: e.message });
  }
}

module.exports = { list, getOne, create, update, remove, bootstrap };
