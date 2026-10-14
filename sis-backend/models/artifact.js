import { DataTypes } from 'sequelize';

// Defines the Artifact model (the application's core content)
const Artifact = (sequelize) => {
    const ArtifactModel = sequelize.define('Artifact', {
        id: {
            type: DataTypes.INTEGER,
            primaryKey: true,
            autoIncrement: true,
        },
        title: {
            type: DataTypes.STRING(255),
            allowNull: false,
        },
        content: {
            type: DataTypes.TEXT,
            allowNull: true,
        },
        owner_id: {
            type: DataTypes.INTEGER,
            allowNull: false,
        },
    }, {
        tableName: 'artifacts',
        timestamps: true,
        createdAt: 'created_at',
        updatedAt: 'updated_at',
    });
    return ArtifactModel;
};

export default Artifact;
