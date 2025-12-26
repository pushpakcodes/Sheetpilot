import mongoose from 'mongoose';
import { MongoMemoryServer } from 'mongodb-memory-server';

const connectDB = async () => {
  try {
    // Try connecting to the provided URI first
    const conn = await mongoose.connect(process.env.MONGO_URI, {
      serverSelectionTimeoutMS: 5000 // Fail fast if local mongo isn't running
    });
    console.log(`MongoDB Connected: ${conn.connection.host}`);
  } catch (error) {
    console.log(`Local MongoDB connection failed: ${error.message}`);
    console.log('Attempting to start in-memory MongoDB...');

    try {
      const mongod = await MongoMemoryServer.create();
      const uri = mongod.getUri();
      const conn = await mongoose.connect(uri);
      console.log(`In-Memory MongoDB Connected: ${conn.connection.host}`);
      // Keep the process alive
    } catch (memError) {
      console.error(`Fatal Error: Could not connect to any MongoDB. ${memError.message}`);
      process.exit(1);
    }
  }
};

export default connectDB;
