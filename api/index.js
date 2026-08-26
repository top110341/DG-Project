// Vercel serverless entry point. vercel.json rewrites every request here explicitly
// (Vercel's automatic "Express" framework-preset routing didn't route nested paths
// like /api/tasks/:id/comments correctly — only single-segment /api/* worked, so this
// makes the routing unambiguous instead of relying on that auto-detection). The
// Express app itself still does all the internal routing (/api/login, /api/tasks,
// static files, etc.), so no route code needed to change.
import app from '../server.js';

export default app;
