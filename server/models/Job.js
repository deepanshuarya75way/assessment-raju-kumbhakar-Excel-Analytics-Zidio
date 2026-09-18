const mongoose = require('mongoose');
const jobSchema = new mongoose.Schema({
  userId:{
    type:mongoose.Schema.Types.ObjectId,
    ref:'User', 
    required:true
  },

  filePath:{ 
    type:String,
     required:true
    },

    status:{
      type:String,
      enum: ['PENDING', 'PROCESSING','COMPLETED','FAILED'],
      default:'PENDING'
    },
    result:{
      type:Object,
      default:null
    },
    errorMessage:{
      type:String,
      default:null
    },
    retryCount:{
      type:Number,
      default:0
    },
    maxRetries:{
      type:Number,
      default:3
    }
    },
    {timestamps:true}
);

module.exports = mongoose.model('Job',jobSchema);