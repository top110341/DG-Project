// Vercel serverless entry point — file-based routing maps every request under
// /api/** to this one function. The Express app itself still does the internal
// routing (/api/login, /api/tasks, etc.), so no route code needed to change.
import app from '../server.js';

export default app;
