// models/Employee.js
import { DataTypes } from 'sequelize';

export default (sequelize) => {
  return sequelize.define("Employee", {
    id: { type: DataTypes.STRING, primaryKey: true },
    full_name: DataTypes.STRING,
    dept: DataTypes.STRING,
    position: DataTypes.STRING,
    basic_salary: DataTypes.DECIMAL(12,2),
    competency: DataTypes.INTEGER,
    // ... other fields as per DDL
  }, { tableName: "employees", timestamps: true });
};