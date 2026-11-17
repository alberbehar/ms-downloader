import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';
import tailwindcss from '@tailwindcss/vite';

import cloudflare from "@astrojs/cloudflare";

const isGhPages = process.env.GITHUB_PAGES === 'true';

export default defineConfig({
  site: isGhPages ? 'https://alberbehar.github.io' : 'https://ms-downloader.example.com',
  base: isGhPages ? '/ms-downloader' : '/',
  integrations: [sitemap()],

  vite: {
    plugins: [tailwindcss()],
  },

  build: {
    inlineStylesheets: 'auto',
  },

  compressHTML: true,
  ...(isGhPages ? {} : { adapter: cloudflare() }),
});