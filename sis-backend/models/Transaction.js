import { DataTypes } from 'sequelize';

// FIX: Export a function that accepts the initialized sequelize instance.
// This prevents the ReferenceError by delaying the 'sequelize.define' call.
const defineTransactionModel = (sequelize) => {
    const Transaction = sequelize.define('Transaction', {
        transactionId: {
            type: DataTypes.UUID,
            defaultValue: DataTypes.UUIDV4,
            primaryKey: true,
            allowNull: false,
        },
        date: {
            type: DataTypes.DATEONLY,
            allowNull: false,
        },
        type: {
            type: DataTypes.STRING,
            allowNull: false,
            // Example: 'Debit', 'Credit'
        },
        account: {
            type: DataTypes.STRING,
            allowNull: false,
            comment: 'Account name (e.g., Cash, Supplies, Income)'
        },
        description: {
            type: DataTypes.STRING,
            allowNull: false,
        },
        amount: {
            type: DataTypes.FLOAT,
            allowNull: false,
            defaultValue: 0.00,
        },
        userId: {
            type: DataTypes.STRING,
            allowNull: true, // Allows for anonymous transactions or if auth is skipped
        }
    }, {
        tableName: 'transactions', // Define the table name explicitly
        modelName: 'Transaction',
    });

    return Transaction;
};

// We still export this function so db.js can call it later.
export default defineTransactionModel;
