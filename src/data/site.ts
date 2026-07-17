// Central configuration & placeholders to confirm.
// All bracketed values are placeholders to be filled in before launch (brief §4).

export const SITE = {
  name: 'Gselltrading',
  domain: 'gselltrading.com',
  url: 'https://gselltrading.com',
  logoParts: { plain: 'GSELL', accent: 'TRADING' },
  tagline:
    'Votre courtier agro-alimentaire de confiance, sur le terrain en Thaïlande.',
  positioning:
    'Votre courtier de confiance sur le terrain en Thaïlande — spécialiste agro-alimentaire, sans stock, à la commission transparente.',
  locale: 'fr',
  // --- Placeholders à confirmer -------------------------------------------
  email: '[EMAIL]',
  phone: '[TELEPHONE]',
  whatsapp: '[TELEPHONE / WHATSAPP]',
  address: {
    locality: 'Phuket',
    region: 'Phuket',
    country: 'TH',
    display: 'Phuket, Thaïlande',
  },
  founderName: '[NOM AFFICHÉ du fondateur]',
  legalEntity: '[Entité légale / LLC]',
  // Réseaux sociaux (laisser vide tant qu'inexistants)
  social: [] as string[],
} as const;

// Whether the site should be indexable by search engines.
// Kept private (noindex + robots Disallow) until launch — flip via env at build.
export const SITE_INDEXABLE =
  import.meta.env.PUBLIC_SITE_INDEXABLE === 'true';

// Main navigation (used by header & footer)
export const NAV = [
  { label: 'Accueil', href: '/' },
  { label: 'Services', href: '/services' },
  { label: 'Rapports', href: '/rapports' },
  { label: 'À propos', href: '/a-propos' },
  { label: 'Contact', href: '/contact' },
] as const;

// The three market reports (used across pages for internal linking)
export const REPORTS = [
  {
    slug: 'panorama-agroalimentaire-thailande',
    title: 'Panorama de l’agro-alimentaire thaïlandais',
    shortTitle: 'Panorama agro-alimentaire',
    summary:
      'Où en est le secteur, quels moteurs, quelles opportunités en 2025-2026.',
  },
  {
    slug: 'marche-riz-thailande',
    title: 'Le marché du riz thaïlandais',
    shortTitle: 'Marché du riz',
    summary:
      'Volumes, prix, variétés et débouchés d’un pilier de l’export thaï.',
  },
  {
    slug: 'marche-thon-conserve-thailande',
    title: 'Le marché du thon en conserve',
    shortTitle: 'Marché du thon en conserve',
    summary:
      'Comment la Thaïlande domine ce marché mondial, et où se trouve la croissance.',
  },
] as const;

// Organization JSON-LD (injected globally in the layout)
export function organizationJsonLd() {
  return {
    '@context': 'https://schema.org',
    '@type': 'Organization',
    name: SITE.name,
    url: SITE.url,
    logo: `${SITE.url}/favicon.svg`,
    description: SITE.positioning,
    address: {
      '@type': 'PostalAddress',
      addressLocality: SITE.address.locality,
      addressRegion: SITE.address.region,
      addressCountry: SITE.address.country,
    },
    ...(SITE.social.length ? { sameAs: SITE.social } : {}),
    contactPoint: {
      '@type': 'ContactPoint',
      contactType: 'sales',
      email: SITE.email,
      areaServed: 'Worldwide',
      availableLanguage: ['French', 'English'],
    },
  };
}

// WebSite JSON-LD (global)
export function webSiteJsonLd() {
  return {
    '@context': 'https://schema.org',
    '@type': 'WebSite',
    name: SITE.name,
    url: SITE.url,
    inLanguage: 'fr',
    publisher: { '@type': 'Organization', name: SITE.name },
  };
}
