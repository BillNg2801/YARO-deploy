const mongoose = require('mongoose');

async function connectDB() {
  const uri = process.env.MONGODB_URI;
  if (!uri) {
    console.warn('MONGODB_URI not set; MongoDB disabled.');
    return null;
  }
  try {
    await mongoose.connect(uri);
    console.log('MongoDB connected:', mongoose.connection.host, '| database:', mongoose.connection.db.databaseName);
    // Ensure "tester" database exists by writing one sample doc (MongoDB creates DB on first write)
    await mongoose.connection.db.collection('samples').updateOne(
      { _id: 'init' },
      { $set: { created: new Date(), note: 'Sample database tester' } },
      { upsert: true }
    );
    return mongoose.connection;
  } catch (err) {
    console.error('MongoDB connection error:', err.message);
    return null;
  }
}

module.exports = { connectDB };
