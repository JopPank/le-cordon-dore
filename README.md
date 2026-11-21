# Le Cordon Doré — Proposition & Site Web Démo

Proposition de refonte de site web moderne, bilingue et mobile-first pour **Le Cordon Doré Restaurant-Bistrot** (58, rue Principale, Vieux-Aylmer, Gatineau).

**Domaine de démonstration :** `https://lecordondore.jopism.com`  
**Site client actuel :** `https://www.lecordondorerestaurantbistrot.ca/`

---

## 📁 Contenu du Projet

* `index.html` — Le site web de démonstration complet :
  * Sticky navigation bilingue (FR / EN instantané) avec bouton d'appel direct
  * Hero section avec badges de réputation (4.8★ Google, viandes 100% certifiées Halal, Chef de Paris)
  * Bande de service en 4 cartes (« Comment voulez-vous manger ? »)
  * Menu interactif par onglets (Indien, Combos Familiaux, Bistro & Grillades, Menu Enfant, Cave Cordon)
  * Encart vedette du **Grand Buffet Indien du mardi soir** et brunchs
  * Section histoire du Chef (25 ans d'art culinaire) et avis clients vérifiés d'Aylmer
  * Bannière service traiteur et cartes-cadeaux
  * Pied de page complet : adresses cliquables, Google Maps intégré, heures d'ouverture détaillées et dissociation appel / texto
* `proposal.html` — Fiche de présentation visuelle imprimable (comparatif direct Wix vs Nouvelle Version)
* `PROPOSAL.md` — Document écrit complet de la proposition commerciale à présenter au propriétaire
* `CNAME` — Configuration DNS pour le sous-domaine `lecordondore.jopism.com`

---

## 🚀 Déploiement vers `lecordondore.jopism.com`

### Option 1 : Cloudflare Pages (Recommandé)
1. Créer un nouveau projet Pages dans votre compte Cloudflare et lier ce dépôt GitHub (ou faire un simple `Direct Upload`).
2. Aucun build command n'est requis (site statique HTML5 pur).
3. Dans **Custom Domains**, ajouter `lecordondore.jopism.com`. Cloudflare gérera le certificat SSL automatiquement.

### Option 2 : GitHub Pages
1. Initialiser git :
   ```bash
   git init
   git add .
   git commit -m "Initial commit - Le Cordon Dore demo site"
   git branch -M main
   git remote add origin <votre-url-github>
   git push -u origin main
   ```
2. Dans GitHub > **Settings > Pages** :
   * Source : `Deploy from a branch` (`main` / root)
   * Custom domain : `lecordondore.jopism.com` (le fichier `CNAME` est déjà présent)
   * Cocher `Enforce HTTPS`

### Option 3 : Netlify ou Vercel
* Glisser-déposer le dossier dans le tableau de bord Netlify ou Vercel.
* Ajouter le domaine personnalisé `lecordondore.jopism.com` avec le CNAME pointant vers votre plateforme.

---

## 🎯 Points Clés de Vente lors du Rendez-vous avec le Propriétaire

1. **Test en direct sur téléphone :** Ouvrez `lecordondore.jopism.com` côte à côte avec leur site Wix actuel sur votre mobile. La différence de vitesse et de lisibilité est immédiate.
2. **Le bouton d'appel direct :** Montrez comment un client peut commander en 1 tap au `819 682-5781`.
3. **Le Buffet du Mardi :** Expliquez comment la section dédiée aide à remplir le restaurant les mardis soirs.
4. **Zéro friction :** Présentez le transfert de nom de domaine sans interruption de service pour leurs clients réguliers.
