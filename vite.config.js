import { defineConfig } from 'vite';

export default defineConfig({
  base: './', // relative paths so it works on any GitHub Pages sub-path
  test: {
    environment: 'node',
    include: ['test/**/*.test.js'],
  },
});
