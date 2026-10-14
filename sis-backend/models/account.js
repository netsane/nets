import { DataTypes } from 'sequelize';

const Account = (sequelize) => {
  const AccountModel = sequelize.define('Account', {
    id: {
      type: DataTypes.INTEGER,
      autoIncrement: true,
      primaryKey: true,
    },
    name: {
      type: DataTypes.STRING,
      allowNull: false,
      unique: true,
      validate: {
        notEmpty: {
          msg: 'Account name cannot be empty.',
        },
      },
    },
    type: { // e.g., 'Asset', 'Liability', 'Equity', 'Revenue', 'Expense'
      type: DataTypes.ENUM('Asset', 'Liability', 'Equity', 'Revenue', 'Expense'),
      allowNull: false,
    },
    description: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
  }, {
    tableName: 'accounts',
    timestamps: true,
  });

  return AccountModel;
};

export default Account;
