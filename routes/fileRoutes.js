const express = require('express');
const fs = require('fs').promises;
const path = require('path');
const mammoth=require('mammoth')
const docx=require('docx');
const {Document,Packer,TextRun}=require('docx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { text } = require('stream/consumers');
const { error } = require('console');
const router = express.Router();

// Helper function to get absolute file path (prevents path traversal attacks)
const getFilePath = (fileName) => path.join(__dirname, '..', 'storage', path.basename(fileName));

// read docx file *******************************************************************************

router.get('/read', async (req, res) => {
  try {
    const fileBuffer = await fs.readFile(getFilePath(req.query.fileName));
    const result=await mammoth.extractRawText({buffer:fileBuffer})

    res.json({ content:result.value });
  } catch (err) {
    res.status(404).json({error:"file not found or its not a DOCX extention"});
  }
});


// Append docx file********************************************************************************************
router.post('/append', async (req, res) => {
  const { fileName, content } = req.body;
  try {
    const filePath=getFilePath(fileName)

    const fileBuffer=await fs.promises.readFile(filePath)

    const zip=new PizZip(fileBuffer)

    const doc=new Docxtemplater(zip)

    const lastContent=doc.getFullText()

    const newContent=lastContent+"\n"+content

    doc.setData({content: newContent})
    doc.render()

    const updatedBuffer=doc.getZip().generate({type:'nodebuffer'})

    await fs.promises.writeFile(filePath,updatedBuffer)
    res.json({ message: 'Content appended successfully' });
    
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

//rename the docx file *******************************************************************************

router.put('/rename', async (req, res) => {
  const { oldName, newName } = req.body;

  if (!oldName || !newName) {
    return res.status(400).json({ error: 'Both old and new file names are required' });
  }

  if(!oldName.endsWith('.docx') || !newName.endsWith('.docx')){
    return res.status(400).json({error:"both old and new names must be a docx extention"})

  }

  const oldFilePath = getFilePath(oldName);
  const newFilePath = getFilePath(newName);

  try {
    // Check if the old file exists
    await fs.access(oldFilePath);

    // Rename the file
    await fs.rename(oldFilePath, newFilePath);
    res.json({ message: 'File renamed successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

//-------------------------------------------------------------------------------------------
router.post('/create-dir', async (req, res) => {
  const { dirName } = req.body;

  if (!dirName) {
    return res.status(400).json({ error: 'Directory name is required' });
  }

  const dirPath = getFilePath(dirName);

  try {
    await fs.mkdir(dirPath, { recursive: true }); // Creates nested directories if needed
    res.json({ message: 'Directory created successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

//------------------------------------------------------------------------------------------------

router.delete('/delete-dir', async (req, res) => {
  const { dirName } = req.query;

  if (!dirName) {
    return res.status(400).json({ error: 'Directory name is required' });
  }

  const dirPath = getFilePath(dirName);

  try {
    await fs.rm(dirPath, { recursive: true, force: true }); // Deletes even if it's not empty
    res.json({ message: 'Directory deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// write docx file****************************************************************************

router.post('/write', async (req, res) => {

    
  try {
    const content=req.body.content
    const doc=new Document({
        sections:[
            {
                properties:{},
                children:[
                    new docx.TextRun(`${content}`),
                    new docx.TextRun({
                        text:"\n",
                        bold:true,
                    }),
                ],
            },
        ],
    });
    const buffer= await Packer.toBuffer(doc)
    await fs.writeFile(getFilePath(req.body.fileName),buffer);
    res.json({ message: 'File written successfully' });
  } catch (err) {
    res.status(500).json({ error:"sorry connt write on this extention"});
  }
});


// delete docx file*********************************************************************************

router.delete('/delete', async (req, res) => {
  try {
    await fs.unlink(getFilePath(req.query.fileName));
    res.json({ message: 'File deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;