const jwt = require('jsonwebtoken');

function authRequired(req, res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (!token) {
    return res.status(401).json({ success: false, message: 'Authentication required' });
  }
  try {
    const payload = jwt.verify(token, process.env.JWT_SECRET || 'unity-sis-dev');
    req.user = payload;
    next();
  } catch (e) {
    return res.status(401).json({ success: false, message: 'Invalid or expired token' });
  }
}

function optionalAuth(req, _res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (token) {
    try {
      req.user = jwt.verify(token, process.env.JWT_SECRET || 'unity-sis-dev');
    } catch (_) {}
  }
  next();
}

function requireModule(...modules) {
  return (req, res, next) => {
    const userMods = req.user?.modules || [];
    const role = String(req.user?.role || '').toLowerCase();
    if (role === 'admin' || role === 'super_admin') return next();
    const ok = modules.some((m) => userMods.map(String).includes(m));
    if (!ok) {
      return res.status(403).json({
        success: false,
        message: `No access to module(s): ${modules.join(', ')}`,
      });
    }
    next();
  };
}

module.exports = { authRequired, optionalAuth, requireModule };
