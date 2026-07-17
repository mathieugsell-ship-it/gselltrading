import type { APIRoute } from 'astro';
import { SITE, SITE_INDEXABLE } from '../data/site';

// Private phase: Disallow all, no sitemap reference (brief §2, §9).
// At launch, set PUBLIC_SITE_INDEXABLE=true → open robots + sitemap reference.
export const GET: APIRoute = () => {
  const body = SITE_INDEXABLE
    ? `User-agent: *
Allow: /

Sitemap: ${SITE.url}/sitemap-index.xml
`
    : `# Site privé — indexation désactivée.
User-agent: *
Disallow: /
`;

  return new Response(body, {
    headers: { 'Content-Type': 'text/plain; charset=utf-8' },
  });
};
