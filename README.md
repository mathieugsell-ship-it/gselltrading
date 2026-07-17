# Gselltrading.com

Site vitrine de **Gselltrading** — courtier agro-alimentaire sur le terrain en
Thaïlande (sourcing, vérification, négociation, suivi ; sans stock, à commission
transparente).

Construit avec **[Astro](https://astro.build)** — sortie 100 % statique,
SEO natif, zéro JS par défaut (îlots interactifs légers pour les micro-animations).

## Développement

```bash
npm install
npm run dev      # serveur local (http://localhost:4321)
npm run build    # build de production → dist/
npm run preview  # prévisualiser le build
```

## Déploiement — Cloudflare Pages

- Commande de build : `npm run build`
- Répertoire de sortie : `dist`
- Node : LTS

## Indexation (privé → public)

Le site est **privé par défaut** : `noindex` global + `robots.txt` en
`Disallow: /`. Toute l'infrastructure SEO (sitemap, JSON-LD, canonicals,
Open Graph) est néanmoins en place.

Pour **passer en public**, il suffit d'un seul flag — la variable
d'environnement de build :

```
PUBLIC_SITE_INDEXABLE=true
```

Cela bascule automatiquement :
- le `<meta name="robots">` en `index, follow` sur les pages publiques ;
- `robots.txt` en mode ouvert + référence du `sitemap-index.xml`.

Ensuite : vérifier le domaine dans Google Search Console + Bing Webmaster
Tools, soumettre le sitemap, contrôler les données structurées (Rich Results
Test) et les Core Web Vitals.

## À confirmer avant lancement

Les valeurs entre crochets sont centralisées dans
[`src/data/site.ts`](src/data/site.ts) :

- `[EMAIL]`, `[TELEPHONE / WHATSAPP]`, `[ADRESSE Phuket]`
- `[NOM AFFICHÉ du fondateur]`, `[Entité légale / LLC]`
- liens réseaux sociaux (`social`)
- clé d'accès du formulaire de contact : `[WEB3FORMS_ACCESS_KEY]` dans
  [`src/pages/contact.astro`](src/pages/contact.astro)
  (voir [web3forms.com](https://web3forms.com))

## Structure

```
src/
├── layouts/       BaseLayout (SEO, header, footer), ReportLayout
├── components/    Header, Footer, Hero, Stat, Faq, ReportCard, CtaSection…
├── pages/         index, services, a-propos, contact, rapports/*, 404, robots.txt
├── data/site.ts   constantes, placeholders, JSON-LD Organization/WebSite
└── styles/        design system global
public/            favicon, images Open Graph
```
