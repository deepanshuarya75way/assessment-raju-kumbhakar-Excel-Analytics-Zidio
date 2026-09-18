const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const bcrypt = require('bcryptjs');
require('dotenv').config();

const app = express();

// Middleware
app.use(cors({
  origin: process.env.CORS_ORIGIN ? process.env.CORS_ORIGIN.split(',') : '*',
  credentials: true
}));
app.use(express.json());

// Cached MongoDB Connection for Serverless Functions
let isConnected = false;

const connectDB = async () => {
  if (isConnected && mongoose.connection.readyState === 1) {
    return;
  }

  const MONGO_URI = process.env.MONGODB_URI;
  if (!MONGO_URI) {
    throw new Error('FATAL ERROR: MONGODB_URI is not defined.');
  }

  try {
    const db = await mongoose.connect(MONGO_URI, {
      bufferCommands: false, // Prevents hanging requests if connection drops
    });
    isConnected = db.connections[0].readyState === 1;
    console.log('MongoDB connected successfully.');

    // Seed superadmin on initial connection
    await seedSuperAdmin();
  } catch (error) {
    console.error('MongoDB connection error:', error);
    throw error;
  }
};

// Middleware to ensure Database Connection before handling routes
app.use(async (req, res, next) => {
  try {
    await connectDB();
    next();
  } catch (error) {
    res.status(500).json({ success: false, message: 'Database connection failed' });
  }
});

// User Schema
const userSchema = new mongoose.Schema({
  email: { type: String, required: true, unique: true },
  password: { type: String, required: true },
  firstName: { type: String, required: true },
  lastName: { type: String, required: true },
  role: { type: String, default: 'user' },
  createdAt: { type: Date, default: Date.now }
});

userSchema.pre('save', async function(next) {
  if (!this.isModified('password')) return next();
  this.password = await bcrypt.hash(this.password, 12);
  next();
});

userSchema.methods.comparePassword = async function(candidatePassword) {
  return bcrypt.compare(candidatePassword, this.password);
};

// Check if model already exists before compiling (prevents OverwriteModelError in Serverless)
const User = mongoose.models.User || mongoose.model('User', userSchema);

// Superadmin Seeder Helper
async function seedSuperAdmin() {
  try {
    const email = process.env.SUPERADMIN_EMAIL || 'superadmin@excelanalytics.app';
    const password = process.env.SUPERADMIN_PASSWORD || 'ChangeMe123!';
    const firstName = 'Super';
    const lastName = 'Admin';

    let adminUser = await User.findOne({ email });
    if (!adminUser) {
      adminUser = new User({ email, password, firstName, lastName, role: 'admin' });
      await adminUser.save();
      console.log(`Seeded superadmin user: ${email}`);
    } else if (adminUser.role !== 'admin') {
      adminUser.role = 'admin';
      await adminUser.save();
      console.log(`Ensured admin role for user: ${email}`);
    }
  } catch (e) {
    console.error('Error seeding superadmin:', e);
  }
}

//Runs on server startup to handle interrupted jobs from previous server restarts
async function recoverIncompleteJobs(){
  const interruptedJobs = await Job.find({ status:'PROCESSING'});
  for(let job of interruptedJobs){
    processJobAsync(job._id); //Re- queue or retry processing safely
  }
}

// Routes
app.get('/', (req, res) => {
  res.json({ message: 'API is running...', timestamp: new Date() });
});

//mount job route
const jobRoutes = require('./routes/jobs');
app.use('/api/jobs',jobRoutes);

// Mount OTP routes
const otpRoutes = require('./routes/otp');
app.use('/api/otp', otpRoutes);

// Mount Admin state routes
const adminRoutes = require('./routes/admin');
const { processJobAsync } = require('./routes/jobs');
app.use('/api/admin', adminRoutes);

// Register endpoint
app.post('/api/auth/register', async (req, res) => {
  try {
    const { email, password, firstName, lastName } = req.body;

    if (!email || !password || !firstName || !lastName) {
      return res.status(400).json({
        success: false,
        message: 'All fields are required'
      });
    }

    const existingUser = await User.findOne({ email });
    if (existingUser) {
      return res.status(400).json({
        success: false,
        message: 'User already exists'
      });
    }

    const user = new User({ email, password, firstName, lastName });
    await user.save();

    res.status(201).json({
      success: true,
      message: 'User registered successfully',
      user: { email, firstName, lastName }
    });

  } catch (error) {
    console.error('Registration error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error'
    });
  }
});

// Login endpoint
app.post('/api/auth/login', async (req, res) => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      return res.status(400).json({
        success: false,
        message: 'Email and password required'
      });
    }

    const user = await User.findOne({ email });
    if (!user) {
      return res.status(401).json({
        success: false,
        message: 'Invalid credentials'
      });
    }

    const isMatch = await user.comparePassword(password);
    if (!isMatch) {
      return res.status(401).json({
        success: false,
        message: 'Invalid credentials'
      });
    }

    res.json({
      success: true,
      message: 'Login successful',
      user: { email: user.email, firstName: user.firstName, lastName: user.lastName }
    });

  } catch (error) {
    console.error('Login error:', error);
    res.status(500).json({
      success: false,
      message: 'Server error'
    });
  }
});

// Start local server if not running on Vercel
if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => {
    console.log(`Server running locally on port ${PORT}`);
  });
}

// Export app for Vercel
module.exports = app;
