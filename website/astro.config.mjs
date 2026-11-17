import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';
import tailwindcss from '@tailwindcss/vite';

import cloudflare from "@astrojs/cloudflare";

export default defineConfig({
  site: 'https://ms-downloader.example.com',
  integrations: [sitemap()],

  vite: {
    plugins: [tailwindcss()],
  },

  build: {
    inlineStylesheets: 'auto',
  },

  compressHTML: true,
  adapter: cloudflare()
});