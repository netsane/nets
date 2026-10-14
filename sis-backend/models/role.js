import { DataTypes } from 'sequelize';

// Defines the Role model (e.g., admin, editor, basic_user)
const Role = (sequelize) => {
    const RoleModel = sequelize.define('Role', {
        id: {
            type: DataTypes.INTEGER,
            primaryKey: true,
            autoIncrement: true,
        },
        name: {
            type: DataTypes.STRING(50),
            allowNull: false,
            unique: true,
        },
    }, {
        tableName: 'roles',
        timestamps: false,
    });
    return RoleModel;
};

export default Role;
