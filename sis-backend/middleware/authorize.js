// sis-backend/middleware/authorize.js
const authorize = (...roles) => {
  return (req, res, next) => {
    // First, ensure the user is authenticated and their role is available
    if (!req.user || !req.user.role) {
      console.error("Authorization middleware: User or role not found in request. Ensure authenticateToken middleware is working correctly.");
      return res.status(401).json({ message: 'Authentication failed or user role missing.' });
    }

    // Check if the user's role is among the allowed roles
    if (!roles.includes(req.user.role)) {
      console.warn(`Authorization: Access denied. User role "${req.user.role}" is not allowed. Required: ${roles.join(', ')}`);
      return res.status(403).json({ message: 'Access denied: Insufficient role.' });
    }

    console.log(`Authorization: User role "${req.user.role}" is allowed for this route.`);
    next(); // User has the required role, proceed to the next middleware/route handler
  };
};

export default authorize;