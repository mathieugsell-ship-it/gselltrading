// @ts-check
import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

// Site is private for now: sitemap is generated but not submitted to engines
// (see robots.txt / brief §9). Flip PUBLIC_SITE_INDEXABLE=true at launch.
export default defineConfig({
  site: 'https://gselltrading.com',
  trailingSlash: 'never',
  integrations: [
    sitemap({
      // Keep noindex utility pages out of the sitemap.
      filter: (page) =>
        !page.includes('/mentions-legales') && !page.includes('/404'),
    }),
  ],
  build: {
    inlineStylesheets: 'auto',
  },
  compressHTML: true,
});
