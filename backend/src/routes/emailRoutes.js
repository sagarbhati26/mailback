import { Router } from 'express';
import multer from 'multer';
import {  sendEmailsFromExcel } from '../controller/emailController.js';



// Set up multer to save uploaded files to the "uploads" directory
const upload = multer({ dest: 'uploads/' });
const router = Router();

// Route for sending bulk emails via JSON body
router.post('/send', sendEmailsFromExcel);

// ✅ Route for uploading Excel and sending personalized emails
router.post('/upload', upload.single('file'), sendEmailsFromExcel);

export default router;