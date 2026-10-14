const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const { query } = require('../config/db');

async function login(req, res) {
  try {
    const username = String(req.body.username || '').trim();
    const password = String(req.body.password || '');
    if (!username || !password) {
      return res.status(400).json({ success: false, message: 'Username and password required' });
    }
    const r = await query(
      `SELECT u.*, ro.name AS role_name, ro.permissions, ro.modules AS role_modules
       FROM users u
       LEFT JOIN roles ro ON ro.id = u.role_id
       WHERE LOWER(u.username) = LOWER($1) AND u.is_active = TRUE`,
      [username]
    );
    const user = r.rows[0];
    if (!user) {
      return res.status(401).json({ success: false, message: 'Invalid username or password' });
    }
    const ok = await bcrypt.compare(password, user.password_hash);
    if (!ok) {
      return res.status(401).json({ success: false, message: 'Invalid username or password' });
    }
    const role = user.role_name || user.role_id || 'employee';
    const modules = user.modules || user.role_modules || ['finance', 'hr', 'material'];
    const token = jwt.sign(
      {
        sub: user.id,
        username: user.username,
        name: user.display_name || user.username,
        role,
        modules,
      },
      process.env.JWT_SECRET || 'unity-sis-dev',
      { expiresIn: process.env.JWT_EXPIRES_IN || '12h' }
    );
    return res.json({
      success: true,
      token,
      user: {
        id: user.id,
        username: user.username,
        name: user.display_name || user.username,
        role,
        modules,
        employee_id: user.employee_id,
      },
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ success: false, message: 'Login failed' });
  }
}

async function me(req, res) {
  return res.json({ success: true, user: req.user });
}

module.exports = { login, me };
