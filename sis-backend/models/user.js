import { DataTypes } from 'sequelize';

// Defines the User model with an association to the Role model
const User = (sequelize) => {
    const UserModel = sequelize.define('User', {
        id: {
            type: DataTypes.INTEGER,
            primaryKey: true,
            autoIncrement: true,
        },
        email: {
            type: DataTypes.STRING(100),
            allowNull: false,
            unique: true,
        },
        password_hash: {
            type: DataTypes.STRING(255),
            allowNull: false,
        },
        first_name: {
            type: DataTypes.STRING(50),
            allowNull: true,
        },
        // foreign key reference to Role
        role_id: {
            type: DataTypes.INTEGER,
            allowNull: false,
        }
    }, {
        tableName: 'users',
        timestamps: false,
    });
    return UserModel;
};

export default User;
