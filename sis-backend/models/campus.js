import { DataTypes } from 'sequelize';

const Campus = (sequelize) => {
  const CampusModel = sequelize.define('Campus', {
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
          msg: 'Campus name cannot be empty.',
        },
      },
    },
    location: {
      type: DataTypes.STRING,
      allowNull: true,
    },
  }, {
    tableName: 'campuses',
    timestamps: true,
  });

  return CampusModel;
};

export default Campus;
