require('dotenv').config();
const express = require('express');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

// Enable CORS so the browser-based Vercel CMS admin page can call this local service
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS, PUT, DELETE");
  
  // Handle preflight options requests
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

// Endpoint called by CMS admin panel when a project is approved/published
app.post('/trigger-resume-update', async (req, res) => {
  try {
    console.log('Received resume update trigger from CMS:', req.body);
    const { name, stack, description } = req.body;

    if (!name || !description) {
      return res.status(400).json({ error: 'Missing name or description in request body.' });
    }

    const pythonScript = path.resolve(__dirname, 'inject_project.py');
    const docPath = path.resolve(__dirname, '../Aby_Kibru_Portfolio.docx');

    // Escape arguments for shell execution
    const titleEscaped = name.replace(/"/g, '\\"');
    const techEscaped = (stack || '').replace(/"/g, '\\"');
    const descEscaped = description.replace(/"/g, '\\"');

    // Run the python injector in-place (input = output = docPath)
    const cmd = `python "${pythonScript}" --title "${titleEscaped}" --tech "${techEscaped}" --desc "${descEscaped}" --input "${docPath}" --output "${docPath}" --column auto --position last`;

    console.log(`Executing: ${cmd}`);
    exec(cmd, (error, stdout, stderr) => {
      if (error) {
        console.error('Python execution error:', stderr || stdout);
        return res.status(500).json({ error: 'Failed to run python injector script.', details: stderr || stdout });
      }
      console.log('Python output:', stdout);
      res.status(200).json({ message: 'Resume updated in-place successfully.', output: stdout });
    });

  } catch (error) {
    console.error('Error triggering update:', error);
    res.status(500).json({ error: 'Failed to trigger update: ' + error.message });
  }
});

// Endpoint to download the updated resume DOCX file directly
app.get('/download-resume', (req, res) => {
  const filePath = path.resolve(__dirname, '../Aby_Kibru_Portfolio.docx');
  console.log(`Download requested for: ${filePath}`);
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Resume template file not found on local PC.' });
  }

  res.download(filePath, 'Aby_Kibru_Portfolio.docx', (err) => {
    if (err) {
      console.error('Download failed:', err);
      if (!res.headersSent) {
        res.status(500).json({ error: 'Failed to download file.' });
      }
    }
  });
});

app.listen(PORT, () => {
  console.log(`Local Resume Updater service listening on port ${PORT}`);
  console.log(`- Webhook: http://localhost:${PORT}/trigger-resume-update`);
  console.log(`- Download: http://localhost:${PORT}/download-resume`);
});
