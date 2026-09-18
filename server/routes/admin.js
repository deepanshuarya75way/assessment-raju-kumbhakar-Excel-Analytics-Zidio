const express = require('express');
const mongoose = require('mongoose');
const AdminState = require('../models/AdminState');
const Job = require('../models/Job');

// Use already-compiled User model from server.js
const User = () => mongoose.model('User');

const router = express.Router();

// Simple header-based protection for server-to-server calls
function requireAdminApiKey(req, res, next) {
  const provided = req.headers['x-admin-auth'];
  const expected = process.env.ADMIN_API_KEY || 'dev-admin-key';
  if (!provided || provided !== expected) {
    return res.status(401).json({ success: false, message: 'Unauthorized' });
  }
  next();
}

async function getSingletonState() {
  let doc = await AdminState.findOne();
  if (!doc) {
    doc = await AdminState.create({});
  }
  return doc;
}

// Read current admin state
router.get('/state', requireAdminApiKey, async (req, res) => {
  try {
    const doc = await getSingletonState();
    res.json({ success: true, state: doc });
  } catch (err) {
    console.error('Admin state GET error:', err);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});

//admin can see system wide progress
router.get('/job-metrics', async (req,res) => {
  const running = await Job.countDocuments({ status: 'PROCESSING'});
  const failed = await Job.countDocuments({status:'FAILED'});
  res.json({runningJobs:running, failedJobs:failed});
})

// Upsert admin state
router.put('/state', requireAdminApiKey, async (req, res) => {
  try {
    const allowed = {
      adminsActive: 'adminsActive',
      adminsPending: 'adminsPending',
      adminsRejected: 'adminsRejected',
      userAdminFlags: 'userAdminFlags',
      sysMaintenance: 'sysMaintenance',
      sysMaintenanceUntil: 'sysMaintenanceUntil',
      sysAlert: 'sysAlert',
      sysUploadLimit: 'sysUploadLimit'
    };

    const update = {};
    for (const k of Object.keys(allowed)) {
      if (k in req.body) update[k] = req.body[k];
    }

    const doc = await getSingletonState();
    Object.assign(doc, update);
    await doc.save();
    res.json({ success: true, state: doc });
  } catch (err) {
    console.error('Admin state PUT error:', err);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});

// Optional utility: ensure email in adminsActive exists as admin user
router.post('/ensure-admin-user', requireAdminApiKey, async (req, res) => {
  try {
    const { email } = req.body || {};
    if (!email) return res.status(400).json({ success: false, message: 'email required' });

    let user = await User().findOne({ email });
    if (!user) {
      const randomPassword = require('crypto').randomBytes(16).toString('hex');
      user = new (User())({ email, password: randomPassword, firstName: 'Admin', lastName: 'User', role: 'admin' });
      await user.save();
    } else if (user.role !== 'admin') {
      user.role = 'admin';
      await user.save();
    }

    res.json({ success: true, user: user.toJSON() });
  } catch (err) {
    console.error('Ensure admin user error:', err);
    res.status(500).json({ success: false, message: 'Server error' });
  }
});

module.exports = router;
