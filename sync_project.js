require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

// Helper to run python injection script
function runPythonInjector(project) {
  return new Promise((resolve, reject) => {
    const pythonScript = path.resolve(__dirname, 'inject_project.py');
    const inputPath = path.resolve(__dirname, '../Aby_Kibru_Portfolio.docx');
    const outputPath = path.resolve(__dirname, '../Aby_Kibru_Portfolio_Updated.docx');
    
    // Escape arguments for the shell execution
    const title = project.title.replace(/"/g, '\\"');
    const tech = project.tech.replace(/"/g, '\\"');
    const desc = project.description.replace(/"/g, '\\"');
    
    const cmd = `python "${pythonScript}" --title "${title}" --tech "${tech}" --desc "${desc}" --input "${inputPath}" --output "${outputPath}"`;
    
    console.log(`Running python injector command: ${cmd}`);
    exec(cmd, (error, stdout, stderr) => {
      if (error) {
        console.error('Python injector failed:', stderr || stdout);
        return reject(error);
      }
      console.log('Python injector stdout:', stdout);
      resolve(stdout);
    });
  });
}

// Helper to call Groq API
async function callGroq(messages, responseFormatJson = true) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey) {
    throw new Error('Missing GROQ_API_KEY in environment variables.');
  }

  const payload = {
    model: 'llama-3.3-70b-versatile',
    messages,
    temperature: 0.1,
  };

  if (responseFormatJson) {
    payload.response_format = { type: 'json_object' };
  }

  const response = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Groq API error (${response.status}): ${errorText}`);
  }

  const data = await response.json();
  const content = data.choices[0].message.content;
  
  if (responseFormatJson) {
    return JSON.parse(content);
  }
  return content;
}

// Fetch repo details from GitHub
async function getRepoData(repoUrlOrName) {
  // Extract owner and repo name
  let cleanPath = repoUrlOrName.replace(/https?:\/\/github\.com\//i, '').replace(/\/$/, '');
  const parts = cleanPath.split('/');
  if (parts.length < 2) {
    throw new Error(`Invalid repository identifier: ${repoUrlOrName}`);
  }
  const owner = parts[parts.length - 2];
  const repo = parts[parts.length - 1];

  const headers = {
    'User-Agent': 'resume-updater',
    'Accept': 'application/vnd.github.v3+json',
  };

  if (process.env.GITHUB_TOKEN) {
    headers['Authorization'] = `token ${process.env.GITHUB_TOKEN}`;
  }

  console.log(`Fetching metadata for ${owner}/${repo} from GitHub...`);
  const repoRes = await fetch(`https://api.github.com/repos/${owner}/${repo}`, { headers });
  if (!repoRes.ok) {
    throw new Error(`Failed to fetch repo ${owner}/${repo}: ${repoRes.statusText}`);
  }
  const repoData = await repoRes.json();

  // Fetch languages
  console.log(`Fetching languages for ${owner}/${repo}...`);
  const langRes = await fetch(`https://api.github.com/repos/${owner}/${repo}/languages`, { headers });
  const languages = langRes.ok ? Object.keys(await langRes.json()).join(', ') : '';

  // Fetch README
  console.log(`Fetching README for ${owner}/${repo}...`);
  let readmeText = '';
  const readmeRes = await fetch(`https://api.github.com/repos/${owner}/${repo}/readme`, { headers });
  if (readmeRes.ok) {
    const readmeJson = await readmeRes.json();
    readmeText = Buffer.from(readmeJson.content, 'base64').toString('utf8');
  } else {
    console.log('README not found via API, trying raw file...');
    const rawReadmeRes = await fetch(`https://raw.githubusercontent.com/${owner}/${repo}/main/README.md`);
    if (rawReadmeRes.ok) {
      readmeText = await rawReadmeRes.text();
    }
  }

  return {
    name: repoData.name,
    fullName: repoData.full_name,
    description: repoData.description || '',
    topics: (repoData.topics || []).join(', '),
    languages,
    readmeText: readmeText.substring(0, 4000), // first 4k chars is enough for Groq
    stars: repoData.stargazers_count,
    githubUrl: repoData.html_url,
  };
}

// Main sync logic
async function syncProject(repoUrlOrName) {
  try {
    const repoData = await getRepoData(repoUrlOrName);
    console.log(`Fetched details for: ${repoData.fullName}`);

    // Call #1: Judge if it's a real project
    console.log('Invoking Groq Judge...');
    const judgePrompt = [
      {
        role: 'system',
        content: `You are an AI assistant for a software developer. Evaluate if the following GitHub repository represents a real coding project (e.g., library, web app, tool, CLI, algorithm) built by the user. 
Answer in JSON format with three fields:
- "is_project" (boolean): true if it is a real portfolio-worthy project, false otherwise.
- "confidence" (number 0-1): your confidence level.
- "reason" (string): brief explanation.

Reject: forks, tutorial exercises, dotfiles, note folders, blank setups, configuration repos, and minor modifications.`,
      },
      {
        role: 'user',
        content: `Repository Name: ${repoData.name}
Description: ${repoData.description}
Topics: ${repoData.topics}
Languages: ${repoData.languages}
README Snippet:
${repoData.readmeText}`,
      },
    ];

    const judgeResult = await callGroq(judgePrompt);
    console.log('Judge Result:', judgeResult);

    if (!judgeResult.is_project) {
      console.log(`Project rejected: ${judgeResult.reason}`);
      return { success: false, status: 'rejected', reason: judgeResult.reason };
    }

    // Call #2: Writer to generate resume/portfolio entry
    console.log('Invoking Groq Writer...');
    const writerPrompt = [
      {
        role: 'system',
        content: `You are a professional resume writer. Write a resume project entry. Return a JSON object with:
- "title" (string): Clean, concise project name.
- "tech" (string): Key technologies used, delimited by " · " (e.g. "React · Node.js · Postgres"). Keep it to the actual main technologies.
- "description" (string): 1-2 sentences of what the project does. Focus on functionality, impact, and technology. DO NOT use generic placeholders or marketing fluff.`,
      },
      {
        role: 'user',
        content: `Repository Name: ${repoData.name}
Description: ${repoData.description}
Languages: ${repoData.languages}
README Snippet:
${repoData.readmeText}`,
      },
    ];

    const projectDetails = await callGroq(writerPrompt);
    console.log('Generated Project Details:', projectDetails);

    // Save to local projects.json (CMS database)
    const projectsDbPath = path.resolve(__dirname, 'projects.json');
    let projects = [];
    if (fs.existsSync(projectsDbPath)) {
      try {
        projects = JSON.parse(fs.readFileSync(projectsDbPath, 'utf8'));
      } catch (e) {
        console.error('Error parsing projects.json, resetting database.', e);
      }
    }

    // Check if repository already exists in CMS to avoid duplicates
    const existingIndex = projects.findIndex(p => p.github_repo === repoData.fullName);
    const newProject = {
      id: existingIndex !== -1 ? projects[existingIndex].id : Date.now(),
      title: projectDetails.title,
      tech: projectDetails.tech,
      description: projectDetails.description,
      github_repo: repoData.fullName,
      github_url: repoData.githubUrl,
      stars: repoData.stars,
      status: 'published', // auto-publish for resume docx updating
      ai_reason: judgeResult.reason,
      synced_at: new Date().toISOString(),
    };

    if (existingIndex !== -1) {
      projects[existingIndex] = newProject;
      console.log('Updated existing project in database.');
    } else {
      projects.push(newProject);
      console.log('Added new project to database.');
    }

    fs.writeFileSync(projectsDbPath, JSON.stringify(projects, null, 2), 'utf8');

    // Run Python injector to update DOCX resume
    console.log('Injecting project into DOCX resume...');
    await runPythonInjector(newProject);

    return { success: true, project: newProject };
  } catch (error) {
    console.error('Sync failed:', error);
    return { success: false, error: error.message };
  }
}

// Allow calling from CLI
if (require.main === module) {
  const repoArg = process.argv[2];
  if (!repoArg) {
    console.error('Usage: node sync_project.js <github_repo_url_or_identifier>');
    process.exit(1);
  }
  syncProject(repoArg).then(res => {
    process.exit(res.success ? 0 : 1);
  });
}

module.exports = { syncProject };
