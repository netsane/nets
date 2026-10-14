// models/TaxConfig.js
import { DataTypes } from 'sequelize';

// Defines the tax law parameters, fetched once at app start or on change.
export default (sequelize) => {
  return sequelize.define("TaxConfig", {
    id: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
    start_date: { type: DataTypes.DATEONLY, allowNull: false },
    end_date: { type: DataTypes.DATEONLY, allowNull: true },
    is_active: { type: DataTypes.BOOLEAN, defaultValue: true },
    // JSONB for the progressive income tax brackets
    pit_brackets: { type: DataTypes.JSONB, allowNull: false,
        comment: '[{"upper_bound": 2000, "rate": 0.00}, ...]'
    },
    pension_employee_rate: { type: DataTypes.DECIMAL(4,3), allowNull: false }, // 0.070
    pension_employer_rate: { type: DataTypes.DECIMAL(4,3), allowNull: false }, // 0.110
    vat_standard_rate: { type: DataTypes.DECIMAL(4,3), allowNull: false }    // 0.150
  }, {
    tableName: "tax_config",
    timestamps: false // Static config, no automatic timestamps needed
  });
};