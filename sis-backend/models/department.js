import { DataTypes } from 'sequelize';

const Department = (sequelize) => {
  const DepartmentModel = sequelize.define('Department', {
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
          msg: 'Department name cannot be empty.',
        },
      },
    },
    campus_id: {
      type: DataTypes.INTEGER,
      allowNull: true,
      references: {
        model: 'campuses', // refers to table name
        key: 'id',
      },
    },
  }, {
    tableName: 'departments',
    timestamps: true,
  });

  return DepartmentModel;
};

export default Department;
