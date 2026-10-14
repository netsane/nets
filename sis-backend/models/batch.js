import { DataTypes } from 'sequelize';

// Export a function that defines and returns the Batch model
export default (sequelize) => sequelize.define('Batch', {
  id: {
    type: DataTypes.INTEGER,
    autoIncrement: true,
    primaryKey: true,
  },
  name: {
    type: DataTypes.STRING,
    allowNull: false,
    unique: true,
  },
  likes: {
    type: DataTypes.INTEGER,
    defaultValue: 0,
  },
  subscriptions: {
    type: DataTypes.INTEGER,
    defaultValue: 0,
  },
}, {
  tableName: 'batches',
  timestamps: true, // Sequelize will manage createdAt and updatedAt
});
