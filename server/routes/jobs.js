const express = require('express');
const router = express.Router();
const multer= require('multer');
const xlsx = require('xlsx');
const Job = require('../models/Job');

const upload = multer({dest:'uploads/'});

async function processJobAsync(jobId){
  const job = await Job.findById(jobId);
  if(!job)return;

  try{
    job.status = 'PROCESSING';
    await job.save();

    //parse excel data
    const workbook = xlsx.readFile(job.filePath);
    const sheetName = workbook.SheetNames[0];
    const parsedData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    //save success status
    job.status = 'COMPLETED';
    job.result = {totalRows:parsedData.length,data:parsedData};
    await job.save();
  } catch(error){
    if(job.retryCount<job.maxRetries){
      job.retryCount+=1;
      await job.save();
      return processJobAsync(jobId);
    }
    //clean fail state with error message
    job.status = 'FAILED';
    job.errorMessage = error.message || 'Failed to process Excel File';
    await job.save();
  }
}

//upload file and create non blocking  job
router.post('/upload',upload.single('file'), async(req,res)=>{
 try{
  const newJob = await Job.create({
    userId:req.user.id,
    filepath:req.file.path,
    status:'PENDING'
  });
  processJobAsync(newJob._id);

  res.status(202).json({ success:true, jobId:newJob._id, status:newJob.status});
 } catch (err) {
  res.status(500).json({
    error:err.message
  })
 }
})

//fetch recent jobs for a user
router.get('/', async(req,res) =>{
  const jobs =await Job.find({userId:req.user.id}).sort({createdAt:-1})
  res.json(jobs);
})

//fetch completed job result 
router.get('/:id', async (req,res) => {
  const job = await Job.findOne({_id:req.params.id, userId:req.user.id});
  if(!job) return res.status(404).json({error:'Job not found'});
  res.json(job);
})

module.exports = { 
  router,
  processJobAsync
};