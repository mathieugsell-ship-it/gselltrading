// Taxonomie produits — SOURCE UNIQUE réutilisée dans tout le site
// (hero, page Services/filières, index Rapports, méga-menu).
// Pour permuter un produit : éditer uniquement ce fichier.

export interface Product {
  key: string;
  nom: string;
  labelTuile: string;
  reportUrl: string;
  icon: string; // id de l'illustration SVG (voir ProductArt.astro)
  accroche: string;
  status: 'live' | 'draft';
  accent: string; // couleur pleine au survol (état relief)
}

export const PRODUCTS: Product[] = [
  {
    key: 'riz',
    nom: 'Riz',
    labelTuile: 'Riz — Hom Mali & blanc',
    reportUrl: '/rapports/marche-riz-thailande',
    icon: 'riz',
    accroche: 'Riz parfumé Hom Mali et riz blanc',
    status: 'live',
    accent: '#d8b76a',
  },
  {
    key: 'thon',
    nom: 'Thon & conserves de la mer',
    labelTuile: 'Thon en conserve',
    reportUrl: '/rapports/marche-thon-conserve-thailande',
    icon: 'thon',
    accroche: 'Leader mondial du thon en conserve',
    status: 'live',
    accent: '#3b8fb0',
  },
  {
    key: 'coco',
    nom: 'Noix de coco',
    labelTuile: 'Lait, crème & eau de coco',
    reportUrl: '/rapports/marche-noix-de-coco-thailande',
    icon: 'coco',
    accroche: 'Lait, crème et eau de coco',
    status: 'draft',
    accent: '#8a5a3c',
  },
  {
    key: 'fruits',
    nom: 'Fruits tropicaux transformés',
    labelTuile: 'Ananas & fruits transformés',
    reportUrl: '/rapports/fruits-tropicaux-transformes-thailande',
    icon: 'fruits',
    accroche: 'Ananas et fruits tropicaux transformés',
    status: 'draft',
    accent: '#e8b84b',
  },
  {
    key: 'poulet',
    nom: 'Poulet transformé (halal)',
    labelTuile: 'Poulet halal transformé',
    reportUrl: '/rapports/poulet-halal-thailande',
    icon: 'poulet',
    accroche: 'Poulet transformé, certifié halal',
    status: 'draft',
    accent: '#e0a96d',
  },
  {
    key: 'petfood',
    nom: 'Aliment pour animaux',
    labelTuile: 'Pet food',
    reportUrl: '/rapports/petfood-thailande',
    icon: 'petfood',
    accroche: 'Aliments pour animaux de compagnie',
    status: 'draft',
    accent: '#2f9e63',
  },
];
