import { DataTypes } from 'sequelize';

// Export a function that defines and returns the Resource model
export default (sequelize) => sequelize.define('Resource', {
  id: {
    type: DataTypes.INTEGER,
    autoIncrement: true,
    primaryKey: true,
  },
  batch_id: {
    type: DataTypes.INTEGER,
    allowNull: false,
    // REMOVE THIS LINE: unique: true,
    references: {
      model: 'batches', // refers to table name
      key: 'id',
    },
  },
  video: {
    type: DataTypes.TEXT,
    defaultValue: "",
  },
  history: {
    type: DataTypes.TEXT,
    defaultValue: "",
  },
  campus_image: {
    type: DataTypes.TEXT,
    defaultValue: "",
  },
  // REMOVED: 'docs' field from Resource model. Documents are now managed solely by the Document model.
}, {
  tableName: 'resources',
  timestamps: true, // Sequelize will manage createdAt and updatedAt
});
