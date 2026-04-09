import { createApp } from './app.js';

const app = createApp();
const PORT = process.env.PORT || 3000;

if (process.env.NODE_ENV !== 'production' || !process.env.VERCEL) {
  app.listen(PORT, () => {
    console.log(`Job Applicator API running on http://localhost:${PORT}`);
  });
}

export default app;
