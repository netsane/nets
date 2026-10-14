const integration = require('../services/integrationService');

async function postMaterial(req, res) {
  try {
    const { source, recordId } = req.body;
    const result = await integration.postMaterialToFinance(source, recordId);
    res.json({ success: true, ...result });
  } catch (e) {
    res.status(400).json({ success: false, message: e.message });
  }
}

async function syncHrPay(req, res) {
  try {
    const result = await integration.syncHrToPayMaster();
    res.json({ success: true, ...result });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message });
  }
}

async function postPayroll(req, res) {
  try {
    const ids = Array.isArray(req.body?.ids) ? req.body.ids : [];
    const result = await integration.postPayrollProcess(ids);
    res.json({ success: true, ...result });
  } catch (e) {
    res.status(500).json({ success: false, message: e.message });
  }
}

module.exports = { postMaterial, syncHrPay, postPayroll };
