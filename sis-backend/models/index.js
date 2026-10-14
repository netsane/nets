import { Sequelize } from 'sequelize';
import sequelize from '../config/db.js'; // Assuming user's file is in config/db.js
import RoleModel from './role.js';
import UserModel from './user.js';
import ArtifactModel from './artifact.js';

// Initialize models
const db = {};
db.sequelize = sequelize;
db.Sequelize = Sequelize;

db.Role = RoleModel(sequelize);
db.User = UserModel(sequelize);
db.Artifact = ArtifactModel(sequelize);

// --- Define Associations ---

// User belongs to Role (Foreign key role_id is in User table)
db.User.belongsTo(db.Role, {
    foreignKey: 'role_id',
    as: 'role' // Alias for fetching: user.getRole()
});
db.Role.hasMany(db.User, {
    foreignKey: 'role_id',
    as: 'users'
});

// Artifact belongs to User (Owner) (Foreign key owner_id is in Artifact table)
db.Artifact.belongsTo(db.User, {
    foreignKey: 'owner_id',
    as: 'owner' // Alias for fetching: artifact.getOwner()
});
db.User.hasMany(db.Artifact, {
    foreignKey: 'owner_id',
    as: 'artifacts'
});


// Optional: Sync models (only if you want Sequelize to create the tables based on the models)
// Use with caution in production. The provided schema.sql is safer for initial setup.
// db.sequelize.sync({ force: false }).then(() => {
//     console.log('Database synchronized.');
// });

export default db;
