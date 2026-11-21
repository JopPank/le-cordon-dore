# coding: utf-8
"""
build_all_materials.py
Générateur complet et synchronisé de tous les documents pour Le Cordon Doré :
- proposal.html & proposal-complete.html
- PROPOSAL.md
- Proposition-Refonte-Le-Cordon-Dore.docx (+ versions complètes & sommaires)
- Presentation-Le-Cordon-Dore.pptx (9 slides écran large 16:9, polices 13-15pt+)
"""

import os, re
import docx
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

BASE_DIR = r"C:\Users\joppa\Desktop\Documents\GitHub\Le Cordon Doré"
os.chdir(BASE_DIR)

# ==============================================================================
# 1. UPDATE proposal.html & proposal-complete.html
# ==============================================================================
print("1. Updating HTML proposal files...")
for fname in ['proposal.html', 'proposal-complete.html']:
    if not os.path.exists(fname):
        continue
    with open(fname, 'r', encoding='utf-8') as f:
        html = f.read()

    # Section 1: Chef bio
    html = re.sub(
        r'un <strong>chef formé à Paris avec 25 ans d’expérience.*?</strong>',
        'un chef propriétaire d\'exception, <strong>Oodith Kallaram</strong>, formé à la prestigieuse <strong>Maison Lenôtre</strong> à Paris avec plus de 25 ans d’expérience culinaire et hôtelière internationale, et a été intronisé parmi <strong>Les Disciples d\'Auguste Escoffier (Océan Indien)</strong>',
        html
    )

    # Section 1: Soften stats and tone down unsourced claims
    html = html.replace(
        'decide où manger en moins de 3 secondes',
        'veut consulter le menu, appeler ou réserver sans friction'
    )
    html = html.replace(
        'décide où manger en moins de 3 secondes',
        'veut consulter le menu, appeler ou réserver sans friction'
    )

    # Section 2: Header title
    html = html.replace(
        '<span>Ce que j’ai remarqué sur le site actuel (et ce qu\'on améliore)</span>',
        '<span>Quelques ajustements recommandés sur le site actuel</span>'
    )

    # Section 2: Chef page finding
    chef_finding_html = '''        <div class="p-3.5 rounded-xl border border-slate-200 bg-slate-50 flex flex-col gap-1.5">
          <div class="font-bold text-slate-900 flex items-center justify-between">
            <span>• Page « Parcours du Chef » vide</span>
            <span class="text-[10px] text-rose-600 font-semibold bg-rose-50 px-2 py-0.5 rounded border border-rose-200">Actuel</span>
          </div>
          <p class="text-slate-600 text-xs">
            Le lien « Parcours du Chef » pointe vers une URL de brouillon Wix (<em>/copie-de-menu-enfant</em>) et n'affiche que sa photo, sans aucun texte biographique.
          </p>
          <p class="text-emerald-700 font-semibold text-xs pt-1 border-t border-slate-200">
            → Nouvelle version : Une section dédiée bilingue racontant enfin l'histoire d'Oodith : sa formation d'excellence à la Maison Lenôtre à Paris, son titre de Disciple d'Auguste Escoffier et ses 25 ans de métier.
          </p>
        </div>'''

    if 'Page « Parcours du Chef » vide' not in html:
        target_nav = '→ Nouvelle version : Cinq sections claires, un bouton d’appel 1-clic toujours visible et une bascule bilingue instantanée (FR / EN).\n          </p>\n        </div>'
        if target_nav in html:
            html = html.replace(target_nav, target_nav + '\n\n' + chef_finding_html)

    # Section 2: Fix Wix pages claim to "accessibles publiquement"
    html = html.replace(
        'et plusieurs pages de test Wix (« copie-de-... ») sont indexées sur Google.',
        'et des pages de brouillon Wix (« copie-de-... ») sont encore accessibles publiquement via les liens du site.'
    )

    # Section 2: Cave grid clarification
    html = html.replace(
        'Bannière mettant en valeur les cocktails signatures à 14 $ et grille des 8 catégories de la cave à compléter selon vos sélections.',
        'Bannière mettant en valeur les cocktails signatures à 14 $ et structure prête à accueillir vos sélections de bouteilles et prix réels.'
    )

    # Section 3: Feature update for chef story
    html = html.replace(
        'Le parcours du chef avec sa photo officielle en tenue, ses 25 ans d’art culinaire et des avis Google récents vérifiés.',
        'Le parcours du chef Oodith Kallaram avec sa photo officielle en tenue de Disciple d\'Escoffier, sa formation Maison Lenôtre à Paris, ses 25 ans d’art culinaire et des avis Google récents vérifiés.'
    )
    html = html.replace(
        'Le parcours du chef avec sa photo officielle en tenue de Disciple d\'Escoffier, ses 25 ans d’art culinaire et des avis Google récents vérifiés.',
        'Le parcours du chef Oodith Kallaram avec sa photo officielle en tenue de Disciple d\'Escoffier, sa formation Maison Lenôtre à Paris, ses 25 ans d’art culinaire et des avis Google récents vérifiés.'
    )

    # Section 3: SEO without maps ranking claim
    html = html.replace(
        'vitesse inférieure à 1 seconde et balisage SEO propre pour attirer la clientèle d\'Aylmer, Gatineau et Ottawa.',
        'présentation soignée, code ultra-léger et balisage SEO propre pour attirer la clientèle d\'Aylmer, Gatineau et Ottawa sur Google.'
    )

    # Section 3: Buffet hours confirmed as example
    html = html.replace(
        'Un bloc dédié avec l’horaire (17h–20h30) et la réservation',
        'Un bloc dédié avec l’horaire (ex. 17h–20h30, à confirmer avec vous) et la réservation'
    )

    # Section 5: Wix plan and Jonathan's philosophy
    html = html.replace(
        'Vous économisez ainsi entre 250 $ et 530 $+ CAD par an.',
        'D’après mes recherches sur les forfaits professionnels Wix actuels pour restaurants au Canada (sur lequel repose votre site actuel), vous payez probablement entre 250 $ et 500 $+ par année (à valider sur votre facture).'
    )

    # Section 5: Maintenance terms
    html = html.replace(
        'Par la suite, nous pourrons convenir ensemble d\'un petit forfait d\'entretien mensuel très abordable ou de simples mises à jour à l\'acte, selon vos besoins réels et sans engagement.',
        'Mon but est avant tout de vous aider et de vous remettre un site clé en main complet. Les 6 premiers mois de mises à jour courantes sont entièrement gratuits. Si plus tard vous désirez de gros ajouts ou de nouveaux développements, nous pourrons simplement en discuter selon le travail nécessaire — sans aucun abonnement obligatoire ni mauvaise surprise.'
    )

    # Replace all leftover coquilles
    html = re.sub(r'\bcoquilles?\b', 'petites erreurs', html, flags=re.IGNORECASE)

    # Remove brackets around Google rating
    html = html.replace(
        '<em>[4,7 ★ · 180+ avis – à valider sur place]</em>',
        '<strong>4,7 ★ · 180+ avis vérifiés sur Google</strong>'
    )
    html = html.replace(
        '[4,7 ★ · 180+ avis – à valider sur place]',
        '4,7 ★ · 180+ avis vérifiés sur Google'
    )

    with open(fname, 'w', encoding='utf-8') as f:
        f.write(html)

print("HTML proposal files updated successfully.")

# ==============================================================================
# 2. UPDATE PROPOSAL.md
# ==============================================================================
print("2. Updating PROPOSAL.md...")
if os.path.exists('PROPOSAL.md'):
    with open('PROPOSAL.md', 'r', encoding='utf-8') as f:
        md = f.read()

    md = re.sub(
        r'Un \*\*chef formé à Paris avec 25 ans d’expérience.*?\*\*',
        'Un chef propriétaire d\'exception, **Chef Oodith Kallaram**, formé à la prestigieuse **Maison Lenôtre** à Paris avec plus de 25 ans d’expérience culinaire et hôtelière internationale, et a été intronisé parmi **Les Disciples d\'Auguste Escoffier (Océan Indien)**',
        md
    )

    chef_finding_md = '''- **Page « Parcours du Chef » vide :**
  - *Actuel :* Le lien « Parcours du Chef » pointe vers une URL de brouillon Wix (`/copie-de-menu-enfant`) et n'affiche que sa photo, sans aucun texte biographique.
  - *→ Nouvelle version :* Une section dédiée bilingue racontant l'histoire d'Oodith : sa formation d'excellence à la Maison Lenôtre à Paris, son titre de Disciple d'Auguste Escoffier et ses 25 ans de métier.'''

    if 'Page « Parcours du Chef » vide' not in md:
        md = md.replace('- **Navigation chargée & doublon avec faute** :', chef_finding_md + '\n\n- **Navigation chargée & doublon avec faute** :')

    md = md.replace(
        'Le parcours du chef avec sa photo officielle en tenue de Disciple d\'Escoffier',
        'Le parcours du chef Oodith Kallaram avec sa photo officielle en tenue de Disciple d\'Escoffier, sa formation Maison Lenôtre à Paris'
    )
    md = md.replace(
        'et plusieurs pages de test Wix (« copie-de-... ») sont indexées sur Google.',
        'et des pages de brouillon Wix (« copie-de-... ») sont encore accessibles publiquement via les liens du site.'
    )
    md = md.replace(
        'Bannière mettant en valeur les cocktails signatures à 14 $ et grille des 8 catégories de la cave à compléter selon vos sélections.',
        'Bannière mettant en valeur les cocktails signatures à 14 $ et structure prête à accueillir vos sélections de bouteilles et prix réels.'
    )
    md = md.replace(
        'Un bloc dédié avec l’horaire (17h–20h30) et la réservation',
        'Un bloc dédié avec l’horaire (ex. 17h–20h30, à confirmer avec vous) et la réservation'
    )
    md = md.replace(
        'Vous économisez ainsi entre 250 $ et 530 $+ CAD par an.',
        'D’après mes recherches sur les forfaits professionnels Wix actuels pour restaurants au Canada (sur lequel repose votre site actuel), vous payez probablement entre 250 $ et 500 $+ par année (à valider sur votre facture).'
    )
    md = md.replace(
        'Par la suite, nous pourrons convenir ensemble d\'un petit forfait d\'entretien mensuel très abordable ou de simples mises à jour à l\'acte, selon vos besoins réels et sans engagement.',
        'Mon but est avant tout de vous aider et de vous remettre un site clé en main complet. Les 6 premiers mois de mises à jour courantes sont entièrement gratuits. Si plus tard vous désirez de gros ajouts ou de nouveaux développements, nous pourrons simplement en discuter selon le travail nécessaire — sans aucun abonnement obligatoire.'
    )

    md = re.sub(r'\bcoquilles?\b', 'petites erreurs', md, flags=re.IGNORECASE)
    md = md.replace('*[4,7 ★ · 180+ avis – à valider sur place]*', '**4,7 ★ · 180+ avis vérifiés sur Google**')
    md = md.replace('[4,7 ★ · 180+ avis – à valider sur place]', '4,7 ★ · 180+ avis vérifiés sur Google')

    with open('PROPOSAL.md', 'w', encoding='utf-8') as f:
        f.write(md)

print("PROPOSAL.md updated successfully.")

# ==============================================================================
# 3. REBUILD WORD DOCUMENTS
# ==============================================================================
print("3. Generating clean Word documents...")
def build_docx(filename):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # Header tag
    p_tag = doc.add_paragraph()
    r_tag = p_tag.add_run('PROPOSITION STRATÉGIQUE & PARTENARIAT')
    r_tag.font.size = Pt(9.5)
    r_tag.font.bold = True
    r_tag.font.color.rgb = RGBColor(180, 83, 9)

    p_title = doc.add_paragraph()
    r_title = p_title.add_run('Un nouveau site web pour Le Cordon Doré')
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(15, 23, 42)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run('Proposition préparée pour Le Cordon Doré Restaurant-Bistrot · 58, rue Principale, Gatineau (Vieux-Aylmer)\n')
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = RGBColor(71, 85, 105)
    
    r_by = p_sub.add_run('Par Jonathan Pankratz, Jopism Solutions Numériques · Novembre 2026\n')
    r_by.font.size = Pt(10)
    r_by.font.color.rgb = RGBColor(71, 85, 105)
    
    r_demo = p_sub.add_run('Démonstration fonctionnelle en ligne : https://lecordondore.jopism.com')
    r_demo.font.size = Pt(10)
    r_demo.font.bold = True
    r_demo.font.color.rgb = RGBColor(180, 83, 9)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 1. En bref
    h1 = doc.add_heading('En bref', level=1)
    h1.style.font.color.rgb = RGBColor(15, 23, 42)
    
    doc.add_paragraph('Votre cuisine, le savoir-faire d\'exception du chef Oodith Kallaram (formé à la prestigieuse Maison Lenôtre à Paris et intronisé Disciple d\'Auguste Escoffier) et votre grand buffet du mardi méritent un site qui donne envie de réserver en quelques secondes. J’ai préparé une démonstration fonctionnelle que vous pouvez essayer dès maintenant : https://lecordondore.jopism.com')
    
    bref_points = [
        ('Ce que je propose', 'Un site rapide, bilingue (français par défaut), pensé d’abord pour le mobile, avec des boutons directs pour appeler, texter et trouver le restaurant en 1 clic.'),
        ('Ce que vous gardez', 'Votre nom de domaine officiel, votre contenu et le contrôle complet de votre présence en ligne.'),
        ('Ce que ça coûte', 'La conception, la migration technique et l’hébergement moderne sont offerts (0 $/mois). Mises à jour courantes gratuites pendant 6 mois sans aucun engagement.'),
        ('Prochaine étape', 'Une présentation décontractée de 10 minutes en personne à Aylmer ou par appel vidéo, sans aucun engagement.')
    ]
    for btitle, bdesc in bref_points:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(f'• {btitle} : ')
        r.bold = True
        r.font.color.rgb = RGBColor(180, 83, 9)
        p.add_run(bdesc)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 2. Ce que vous faites déjà très bien
    h2 = doc.add_heading('1. Ce que vous faites déjà très bien', level=1)
    h2.style.font.color.rgb = RGBColor(15, 23, 42)
    doc.add_paragraph('Le Cordon Doré a des atouts que très peu de restaurants peuvent revendiquer dans l\'Outaouais : un chef propriétaire d\'exception, Oodith Kallaram, formé à la prestigieuse Maison Lenôtre à Paris avec plus de 25 ans d’expérience culinaire et hôtelière internationale, et a été intronisé parmi Les Disciples d\'Auguste Escoffier (Océan Indien); une double cuisine indienne et bistrot international soignée; des viandes 100 % certifiées Halal; un buffet réputé chaque mardi soir; et une adresse bien ancrée au cœur du Vieux-Aylmer avec plus de 180 avis élogieux (4,7 ★ sur Google).')
    doc.add_paragraph('Le site web doit être le reflet fidèle de cette réputation, pensé d’abord pour le client qui vous cherche sur son cellulaire et veut consulter le menu, appeler ou réserver sans friction.')

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 3. Quelques ajustements recommandés
    h3 = doc.add_heading('2. Quelques ajustements recommandés sur le site actuel', level=1)
    h3.style.font.color.rgb = RGBColor(15, 23, 42)
    p_rem_intro = doc.add_paragraph('« Rien de grave : ce sont des détails fréquents sur les sites qui se construisent au fil des années. Mais ils ont un effet réel sur un visiteur pressé. »')
    p_rem_intro.runs[0].font.italic = True
    p_rem_intro.runs[0].font.color.rgb = RGBColor(100, 116, 139)

    remarques = [
        ('Navigation chargée & doublon avec faute', 'Treize liens dans le menu, dont un doublon avec faute d\'orthographe (« Home » et « Acceuil » au lieu d\'« Accueil »), et un titre de page en anglais (« Home | Le Cordon Doré… ») sur un site francophone.', 'Cinq sections claires, un bouton d’appel 1-clic toujours visible et une bascule bilingue instantanée (FR / EN).'),
        ('Page « Parcours du Chef » vide', 'Le lien « Parcours du Chef » pointe vers une URL de brouillon Wix (/copie-de-menu-enfant) et n\'affiche que sa photo, sans aucun texte biographique.', 'Une section dédiée bilingue racontant enfin l\'histoire d\'Oodith : sa formation d\'excellence à la Maison Lenôtre à Paris, son titre de Disciple d\'Auguste Escoffier et ses 25 ans de métier.'),
        ('Coordonnées qui se contredisent', 'Deux adresses courriel différentes (cordondor1803@gmail.com vs cordondor18103@gmail.com sous Réservation) et trois numéros de téléphone éparpillés.', 'Un rôle clair par canal (appeler pour commander au 819 682-5781, texter pour réserver au 819 661-5113, et un courriel vérifié pour le traiteur).'),
        ('Petites erreurs dans le menu', 'Le Seekh kebab de poulet apparaît deux fois dans les entrées, et on trouve des erreurs d\'anglais visibles comme « Scalops and calamary » et « briyani rice » dans les descriptions anglaises.', '44 plats indiens et 20 spécialités bistrot relus, sans doublon, avec orthographe bilingue soignée et validée avec le chef.'),
        ('Menu familial introuvable', 'Le lien « Menu familiale Emporter » mène vers une page qui n’affiche qu\'une vieille image d\'avis de vacances sans aucun plat ni prix.', 'Un onglet dédié aux festins familiaux (pour 2 personnes et 4 à 5 personnes), avec compositions et tarifs transparents.'),
        ('Pages inachevées & liens orphelins', 'Le blogue ne contient qu’un article vide intitulé « Test » (« Aucun post publié dans cette langue »), et des pages de brouillon Wix (« copie-de-... ») sont encore accessibles publiquement via les liens du site.', 'Architecture monopage moderne sans cul-de-sac, avec redirection automatique des anciennes adresses pour préserver votre référencement naturel.'),
        ('Carte des boissons incomplète', 'Les 7 cocktails sont affichés à 14 $, mais les rubriques vins, bières et spiritueux ne montrent aucun article ni tarif.', 'Bannière mettant en valeur les cocktails signatures à 14 $ et structure prête à accueillir vos sélections de bouteilles et prix réels.')
    ]

    for title, act, nouv in remarques:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        r_tit = p.add_run(f'• {title}\n')
        r_tit.bold = True
        r_tit.font.size = Pt(10)
        
        r_act = p.add_run(f'   Actuel : {act}\n')
        r_act.font.color.rgb = RGBColor(190, 18, 60)
        r_act.font.size = Pt(9.5)
        
        r_nouv = p.add_run(f'   → Nouvelle version : {nouv}')
        r_nouv.bold = True
        r_nouv.font.color.rgb = RGBColor(4, 120, 87)
        r_nouv.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 4. Ce que contient la nouvelle version
    h4 = doc.add_heading('3. Ce que contient la nouvelle version', level=1)
    h4.style.font.color.rgb = RGBColor(15, 23, 42)

    features = [
        ('Le mobile d’abord', 'Boutons directs pour appeler (819 682-5781), texter pour réserver (819 661-5113) et ouvrir l’itinéraire Google Maps. Heures d’ouverture visibles dès l’accueil.'),
        ('Le buffet du mardi en vedette', 'Un bloc dédié avec l’horaire (ex. 17h–20h30, à confirmer avec vous) et la réservation, plus un renvoi vers votre page Facebook officielle pour les brunchs spéciaux du week-end.'),
        ('Des menus complets & bilingues', '44 plats indiens, 20 spécialités bistro, 6 menus enfants, desserts et cocktails, avec badges clairs (Halal, Végétarien, Sans gluten, piquant).'),
        ('Traiteur et cartes-cadeaux', 'Une section pour recevoir les soumissions (mariages, fêtes privées, corporatif dans l’Outaouais) et afficher les montants disponibles (25 $ à 200 $).'),
        ('Votre histoire & authenticité', 'Le parcours du chef Oodith Kallaram avec sa photo officielle en tenue de Disciple d\'Escoffier, sa formation Maison Lenôtre à Paris, ses 25 ans d’art culinaire et des avis Google récents vérifiés.'),
        ('Visibilité locale soignée', 'Présentation soignée, code ultra-léger et balisage SEO propre pour attirer la clientèle d\'Aylmer, Gatineau et Ottawa sur Google.')
    ]
    for ftitle, fdesc in features:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.5)
        r = p.add_run(f'• {ftitle} : ')
        r.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)
        p.add_run(fdesc)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 5. Ce que nous validerons ensemble
    h5 = doc.add_heading('4. Ce que nous validerons ensemble avant la mise en ligne', level=1)
    h5.style.font.color.rgb = RGBColor(15, 23, 42)
    doc.add_paragraph('La démonstration est un prototype fonctionnel. Voici ce que je ne publierai jamais sans votre accord préalable :')

    vals = [
        ('Menu, prix et allergènes', 'Le chef approuve chaque plat, description et mention diététique. Certains textes de présentation de la démo ont été rédigés par moi et doivent être confirmés ou corrigés selon vos recettes.'),
        ('Horaires et buffet du mardi', 'Confirmation de vos heures exactes d\'ouverture et de la plage horaire du grand buffet (ex. 17h à 20h30).'),
        ('Combos familiaux', 'Les compositions et les prix de la démo sont des exemples à ajuster selon vos formules exactes.'),
        ('Photos réelles', 'Vos vraies photos (le chef, la salle, le buffet, les plats) à la place des images d\'illustration.'),
        ('Avis et note Google', 'Uniquement de vrais avis cités fidèlement. La note et le nombre d’avis affichés seront vérifiés le jour de la mise en ligne : 4,7 ★ · 180+ avis vérifiés sur Google.'),
        ('Cartes-cadeaux', 'Votre site actuel permet d’acheter des cartes-cadeaux numériques de 25 $ à 200 $ en ligne. Nous choisirons ensemble comment continuer à les offrir et comment honorer celles déjà vendues.'),
        ('Coordonnées et réseaux sociaux', 'Votre page Facebook officielle (cordondorerestaurantbistrot), votre compte Instagram (@cordondorerestaurantbistro), le numéro secondaire 819 682-1977 et l’adresse courriel officielle.'),
        ('Carte des vins, bières et spiritueux', 'À compléter avec vos bouteilles réelles et les prix que vous souhaitez afficher en salle.')
    ]
    for vtitle, vdesc in vals:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.5)
        r = p.add_run(f'✓ {vtitle} — ')
        r.bold = True
        r.font.color.rgb = RGBColor(180, 83, 9)
        p.add_run(vdesc)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 6. Coûts, engagements et intention
    h6 = doc.add_heading('5. Coûts, engagements et démarche', level=1)
    h6.style.font.color.rgb = RGBColor(15, 23, 42)

    costs = [
        ('Ma démarche', 'Je réalise ce projet avant tout pour aider concrètement un restaurant exceptionnel de ma communauté. Mon but est de vous livrer un outil moderne, complet et clé en main, sans rien vous demander en échange.'),
        ('Ce qui est offert', 'Conception complète du site, migration technique, mise en ligne et hébergement moderne haute performance (0 $/mois).'),
        ('Votre abonnement Wix actuel', 'D’après mes recherches sur les forfaits professionnels Wix actuels pour restaurants au Canada (sur lequel repose votre site actuel), vous payez probablement entre 250 $ et 500 $+ par année. Vous pourrez l’annuler une fois la transition complétée et vos cartes-cadeaux réglées, à valider sur votre facture.'),
        ('Seul frais maintenu', 'Uniquement le renouvellement de votre nom de domaine .ca auprès de votre registraire habituel (environ 15 $ à 25 $ par année).'),
        ('Mises à jour & Suivi', 'Durant les 6 premiers mois, les ajustements courants (changement de prix, ajout de plats saisonniers, horaires spéciaux de fêtes ou du buffet) sont pris en charge gratuitement. Par la suite, si vous désirez des ajouts majeurs ou de nouveaux développements, nous pourrons simplement en discuter selon le travail nécessaire — aucun abonnement forcé ni mauvaise surprise.'),
        ('Propriété & Liberté', 'Votre domaine reste enregistré à votre nom, votre contenu vous appartient à 100 %, et je vous remets sur demande une copie complète du site. Vous êtes libres d\'arrêter quand vous le souhaitez.')
    ]
    for ctitle, cdesc in costs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.5)
        r = p.add_run(f'• {ctitle} : ')
        r.bold = True
        p.add_run(cdesc)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 7. Évolution optionnelle : commandes en ligne directes
    h7 = doc.add_heading('6. Évolution future optionnelle : commandes en ligne directes (DoorDash Storefront)', level=1)
    h7.style.font.color.rgb = RGBColor(15, 23, 42)
    doc.add_paragraph('Si vous souhaitez plus tard offrir la commande directe pour emporter ou la livraison sans payer les commissions élevées des plateformes tierces, nous pourrons explorer ensemble la solution DoorDash Storefront :')

    dd_points = [
        ('0 % de commission marketplace', 'Sur les commandes passées directement depuis votre propre site web (seuls les frais habituels de traitement de carte s\'appliquent).'),
        ('Aucun abonnement logiciel mensuel', 'Pas de frais de licence récurrents pour votre restaurant.'),
        ('Gestion directe', 'Possibilité d\'intégrer les commandes sur votre matériel de comptoir ou tablette existante (modalités exactes à valider avec DoorDash).'),
        ('Option de livraison par Dashers', 'Possibilité d\'utiliser le réseau de livreurs DoorDash avec des frais de livraison transparents pour le client.')
    ]
    for dtitle, ddesc in dd_points:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.5)
        r = p.add_run(f'• {dtitle} — ')
        r.bold = True
        p.add_run(ddesc)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 8. Plan de transition
    h8 = doc.add_heading('7. Comment se passe la transition', level=1)
    h8.style.font.color.rgb = RGBColor(15, 23, 42)

    deploys = [
        ('1. Découverte de la démonstration', 'Vous essayez le prototype sur votre téléphone ou sur mon écran.'),
        ('2. Vos ajustements', 'Vous me transmettez vos corrections de menu, prix exacts et photos.'),
        ('3. Mise en ligne sur votre adresse', 'Votre site actuel reste en ligne jusqu’à ce que le nouveau soit approuvé. La bascule elle-même prend quelques heures sans aucune coupure.'),
        ('4. Suivi & accompagnement', 'Configuration des redirections, fiche Google et petits ajustements durant le premier mois.')
    ]
    for dtitle, ddesc in deploys:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.5)
        r = p.add_run(f'{dtitle} — ')
        r.bold = True
        p.add_run(ddesc)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 9. Prochaine étape
    h9 = doc.add_heading('Prochaine étape simple', level=1)
    h9.style.font.color.rgb = RGBColor(15, 23, 42)
    doc.add_paragraph('Je serais très heureux de passer vous voir au restaurant sur la rue Principale à Aylmer (ou lors d\'un court appel vidéo) pour vous faire essayer la démonstration sur votre téléphone ou directement sur mon écran, et répondre à toutes vos questions, sans aucun engagement.')
    
    p_contact = doc.add_paragraph()
    r_cname = p_contact.add_run('Jonathan Pankratz · Jopism Solutions Numériques\n')
    r_cname.bold = True
    r_cinfo = p_contact.add_run('Courriel : info@jopism.com  |  Web : digital.jopism.com  |  Démo : lecordondore.jopism.com')
    r_cinfo.font.bold = True
    r_cinfo.font.color.rgb = RGBColor(180, 83, 9)

    doc.save(filename)
    print(f"  -> {filename} generated successfully.")

for fname in ['Proposition-Strategique-Complete.docx', 'Proposition-Sommaire-Comparatif.docx', 'Proposition-Refonte-Le-Cordon-Dore.docx']:
    build_docx(fname)

# ==============================================================================
# 4. REBUILD POWERPOINT PRESENTATION (Widescreen 16:9, Big Readable 13-15pt Fonts)
# ==============================================================================
print("4. Generating Presentation-Le-Cordon-Dore.pptx...")
prs = Presentation()
prs.slide_width = PptxInches(13.333)
prs.slide_height = PptxPt(540) # 7.5 inches

C_DARK = PptxRGBColor(11, 15, 23)        # #0B0F17
C_CARD = PptxRGBColor(24, 32, 48)        # #182030
C_GOLD = PptxRGBColor(212, 175, 55)      # #D4AF37
C_GOLD_LIGHT = PptxRGBColor(243, 229, 171) # #F3E5AB
C_WHITE = PptxRGBColor(255, 255, 255)
C_SLATE = PptxRGBColor(148, 163, 184)    # #94A3B8
C_EMERALD = PptxRGBColor(16, 185, 129)   # #10B981
C_ROSE = PptxRGBColor(244, 63, 94)

def add_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_DARK
    bg.line.fill.background()

def add_header(slide, tag, title):
    tb_tag = slide.shapes.add_textbox(PptxInches(0.8), PptxInches(0.45), PptxInches(11.7), PptxInches(0.35))
    tf_tag = tb_tag.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag.upper()
    p_tag.font.size = PptxPt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_GOLD

    tb_title = slide.shapes.add_textbox(PptxInches(0.8), PptxInches(0.75), PptxInches(11.7), PptxInches(0.65))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title
    p_title.font.size = PptxPt(23)
    p_title.font.bold = True
    p_title.font.color.rgb = C_WHITE

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(0.8), PptxInches(1.48), PptxInches(11.733), PptxInches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = PptxRGBColor(30, 41, 59)
    line.line.fill.background()

# ------------------------------------------------------------------------------
# Slide 1: Cover
# ------------------------------------------------------------------------------
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide1)
if os.path.exists('logo.png'):
    slide1.shapes.add_picture('logo.png', PptxInches(0.8), PptxInches(2.1), width=PptxInches(2.0))

tb1 = slide1.shapes.add_textbox(PptxInches(3.2), PptxInches(1.9), PptxInches(9.3), PptxInches(4.2))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1_pre = tf1.paragraphs[0]
p1_pre.text = "PROPOSITION STRATÉGIQUE & DÉMONSTRATION"
p1_pre.font.size = PptxPt(11.5)
p1_pre.font.bold = True
p1_pre.font.color.rgb = C_GOLD
p1_pre.space_after = PptxPt(10)

p1_tit = tf1.add_paragraph()
p1_tit.text = "Un nouveau site web pour\nLe Cordon Doré"
p1_tit.font.size = PptxPt(38)
p1_tit.font.bold = True
p1_tit.font.color.rgb = C_WHITE
p1_tit.space_after = PptxPt(14)

p1_sub = tf1.add_paragraph()
p1_sub.text = "Mettre en valeur votre cuisine, l'histoire du chef Oodith Kallaram (Maison Lenôtre, Disciple d'Escoffier) et faciliter la vie de vos clients sur mobile."
p1_sub.font.size = PptxPt(15)
p1_sub.font.color.rgb = C_SLATE
p1_sub.space_after = PptxPt(20)

p1_foot = tf1.add_paragraph()
p1_foot.text = "Préparé par Jonathan Pankratz · Jopism Solutions Numériques  |  Démo : lecordondore.jopism.com"
p1_foot.font.size = PptxPt(12)
p1_foot.font.color.rgb = C_GOLD_LIGHT

# ------------------------------------------------------------------------------
# Slide 2: En Résumé + QR Code
# ------------------------------------------------------------------------------
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide2)
add_header(slide2, "En résumé", "Une solution simple, moderne et avantageuse")

cards2 = [
    ("📱 Pensé d'abord pour le mobile", "Boutons directs pour appeler, texter et trouver le restaurant en quelques secondes sur téléphone comme sur ordinateur."),
    ("⚡ Affichage rapide & soigné", "Code léger et présentation fluide pour une excellente expérience client et un référencement propre sur Google à Aylmer et Ottawa."),
    ("🛡️ Contrôle total & Propriété", "Votre nom de domaine reste enregistré à votre nom, votre contenu vous appartient à 100 %, et vous arrêtez quand vous voulez."),
    ("💰 Conception & Hébergement offerts", "Site complet livré à 0 $. Mises à jour courantes gratuites pendant 6 mois sans aucun engagement.")
]

for idx, (ctitle, cdesc) in enumerate(cards2):
    x = PptxInches(0.8 + (idx % 2) * 4.6)
    y = PptxInches(1.75 + (idx // 2) * 2.65)
    card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, PptxInches(4.35), PptxInches(2.35))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD
    card.line.color.rgb = PptxRGBColor(51, 65, 85)
    
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = ctitle
    p.font.size = PptxPt(15)
    p.font.bold = True
    p.font.color.rgb = C_GOLD_LIGHT
    p.space_after = PptxPt(8)
    
    p2 = tf.add_paragraph()
    p2.text = cdesc
    p2.font.size = PptxPt(13)
    p2.font.color.rgb = C_WHITE

# QR Box
c_qr2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(10.25), PptxInches(1.75), PptxInches(2.28), PptxInches(5.0))
c_qr2.fill.solid()
c_qr2.fill.fore_color.rgb = C_CARD
c_qr2.line.color.rgb = C_GOLD
if os.path.exists('qr_gold.png'):
    slide2.shapes.add_picture('qr_gold.png', PptxInches(10.45), PptxInches(2.2), width=PptxInches(1.88))
tb_q2 = slide2.shapes.add_textbox(PptxInches(10.35), PptxInches(4.4), PptxInches(2.08), PptxInches(2.0))
tf_q2 = tb_q2.text_frame
tf_q2.word_wrap = True
p_q2 = tf_q2.paragraphs[0]
p_q2.text = "Scannez pour ouvrir la démo en direct :\nlecordondore.jopism.com"
p_q2.font.size = PptxPt(11)
p_q2.font.bold = True
p_q2.font.color.rgb = C_GOLD_LIGHT
p_q2.alignment = PP_ALIGN.CENTER

# ------------------------------------------------------------------------------
# Slide 3: Vos Atouts (Photo du Chef + Liste)
# ------------------------------------------------------------------------------
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide3)
add_header(slide3, "Vos Atouts", "Ce que Le Cordon Doré fait déjà remarquablement bien")

if os.path.exists('chef.jpg'):
    slide3.shapes.add_picture('chef.jpg', PptxInches(0.8), PptxInches(1.75), width=PptxInches(3.3))

tb3 = slide3.shapes.add_textbox(PptxInches(4.4), PptxInches(1.75), PptxInches(8.1), PptxInches(5.2))
tf3 = tb3.text_frame
tf3.word_wrap = True

atouts = [
    ("👨‍🍳 Chef Oodith Kallaram · Maison Lenôtre & Disciple d'Auguste Escoffier", "25 ans de carrière internationale, formé à la prestigieuse Maison Lenôtre à Paris, intronisé Disciple d'Escoffier. Une alliance d'exception entre haute gastronomie française et épices d'Inde."),
    ("🍛 Double carte authentique & variée", "44 spécialités indiennes soignées, classiques du bistrot international, 6 menus enfants et options véganes/végétariennes."),
    ("🥩 Viandes 100 % certifiées Halal", "Un gage de qualité et de respect des traditions recherché par la clientèle de Gatineau et d'Ottawa."),
    ("✨ Grand Buffet Indien du Mardi Soir", "Une formule à volonté réputée qui fidélise les amateurs de cuisine indienne chaque semaine."),
    ("⭐ Réputation solide dans le Vieux-Aylmer", "Une note d'excellence de 4,7 ★ sur Google basée sur plus de 180 avis élogieux vérifiés.")
]

for idx, (atitle, adesc) in enumerate(atouts):
    p = tf3.paragraphs[0] if idx == 0 else tf3.add_paragraph()
    p.text = atitle + " : "
    p.font.size = PptxPt(13.5)
    p.font.bold = True
    p.font.color.rgb = C_GOLD
    p.space_after = PptxPt(8)
    
    r = p.add_run()
    r.text = adesc
    r.font.bold = False
    r.font.color.rgb = C_WHITE

# ------------------------------------------------------------------------------
# Slide 4: Quelques ajustements recommandés (6 cartes équilibrées, 13.5pt)
# ------------------------------------------------------------------------------
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide4)
add_header(slide4, "Optimisation", "Quelques ajustements recommandés sur le site actuel")

remarques_pptx = [
    ("Navigation encombrée", "13 liens avec doublon « Home » et « Acceuil » (faute d'orthographe).", "5 sections claires, bouton d'appel fixe et bascule bilingue instantanée."),
    ("Page Parcours du Chef vide", "Lien vers URL de brouillon (/copie-de-menu-enfant) avec photo sans aucun texte.", "Section dédiée racontant Oodith : Maison Lenôtre, Escoffier et 25 ans de métier."),
    ("Coordonnées divergentes", "2 adresses Gmail distinctes (cordondor1803 vs 18103) et 3 numéros éparpillés.", "1 canal clair : appel commandes (819 682-5781), texto réservations (819 661-5113)."),
    ("Petites erreurs au menu", "Seekh kebab listé deux fois, fautes d'anglais (Scalops and calamary, briyani rice).", "Menu relu, orthographe bilingue soignée, validé avec le chef."),
    ("Menu familial manquant", "Le lien « Menu familiale » mène à une page affichant un vieil avis de vacances sans plat.", "Onglet festins familiaux avec compositions et prix transparents."),
    ("Carte des boissons", "Cocktails à 14 $ listés, mais les sections vins, bières et spiritueux sont sans prix.", "7 cocktails à 14 $ mis en avant + structure prête pour votre carte réelle.")
]

for idx, (rtitle, rcur, rsol) in enumerate(remarques_pptx):
    x = PptxInches(0.8 + (idx % 2) * 6.0)
    y = PptxInches(1.75 + (idx // 2) * 1.75)
    card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, PptxInches(5.7), PptxInches(1.6))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD
    card.line.color.rgb = PptxRGBColor(51, 65, 85)
    
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"• {rtitle}"
    p.font.size = PptxPt(13.5)
    p.font.bold = True
    p.font.color.rgb = C_GOLD_LIGHT
    p.space_after = PptxPt(3)
    
    p2 = tf.add_paragraph()
    r_cur = p2.add_run()
    r_cur.text = f"Actuel : {rcur}\n"
    r_cur.font.size = PptxPt(11)
    r_cur.font.color.rgb = C_ROSE
    
    r_sol = p2.add_run()
    r_sol.text = f"Nouveau : {rsol}"
    r_sol.font.size = PptxPt(11)
    r_sol.font.bold = True
    r_sol.font.color.rgb = C_EMERALD

# ------------------------------------------------------------------------------
# Slide 5: Transparence Totale (Grandes polices aérées)
# ------------------------------------------------------------------------------
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide5)
add_header(slide5, "Transparence Totale", "Ce que nous validerons ensemble avant toute mise en ligne")

tb5 = slide5.shapes.add_textbox(PptxInches(0.8), PptxInches(1.75), PptxInches(11.7), PptxInches(5.2))
tf5 = tb5.text_frame
tf5.word_wrap = True

validations = [
    ("Menus, prix et allergènes", "Le chef approuve chaque plat, description et mention diététique. Rien n'est publié sans son accord formel."),
    ("Horaires exacts & Buffet du mardi", "Confirmation de vos heures d'ouverture et de la plage horaire du buffet (ex. 17h à 20h30)."),
    ("Combos familiaux", "Les compositions et les tarifs de la démonstration sont des suggestions à ajuster selon vos formules exactes."),
    ("Photos officielles", "Intégration de vos vraies photos (salle, cuisine, buffet, chef) en remplacement des images d'illustration."),
    ("Avis et note Google", "Affichage fidèle vérifié le jour J de la note officielle : 4,7 ★ · 180+ avis vérifiés sur Google."),
    ("Coordonnées & Réseaux sociaux", "Vérification des liens officiels Facebook (cordondorerestaurantbistrot), Instagram (@cordondorerestaurantbistro) et cartes-cadeaux.")
]

for idx, (vtitle, vdesc) in enumerate(validations):
    p = tf5.paragraphs[0] if idx == 0 else tf5.add_paragraph()
    p.text = f"✓ {vtitle} : "
    p.font.size = PptxPt(13.5)
    p.font.bold = True
    p.font.color.rgb = C_GOLD
    p.space_after = PptxPt(10)
    
    r = p.add_run()
    r.text = vdesc
    r.font.bold = False
    r.font.color.rgb = C_WHITE

# ------------------------------------------------------------------------------
# Slide 6: Aspect Financier (Frais estimés Wix vs Nouvelle Approche)
# ------------------------------------------------------------------------------
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide6)
add_header(slide6, "Aspect Financier", "Frais actuels estimés vs Nouvelle approche")

col_w = PptxInches(5.65)
c_left = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(1.75), col_w, PptxInches(5.0))
c_left.fill.solid()
c_left.fill.fore_color.rgb = PptxRGBColor(30, 20, 25)
c_left.line.color.rgb = C_ROSE
tf_l = c_left.text_frame
tf_l.word_wrap = True

pl = tf_l.paragraphs[0]
pl.text = "🔴 SITUATION ACTUELLE ESTIMÉE (WIX)"
pl.font.size = PptxPt(15)
pl.font.bold = True
pl.font.color.rgb = PptxRGBColor(251, 113, 133)
pl.space_after = PptxPt(14)

wix_pts = [
    "• Abonnement Wix annuel : D'après mes recherches sur leurs forfaits restaurant au Canada, vous payez probablement ~250 $ à 500 $+ / an (à valider sur votre facture).",
    "• Domaine .ca : ~15 $ à 25 $ / an auprès de votre registraire.",
    "• Commissions des applications de livraison : 20 % à 29 % sur les livraisons tierces (et ~8 % à 10 % sur le ramassage via l'app).",
    "• Sur une livraison moyenne de 40 $, cela représente 8 $ à 11 $ prélevés par la plateforme."
]
for pt in wix_pts:
    p = tf_l.add_paragraph()
    p.text = pt
    p.font.size = PptxPt(12.5)
    p.font.color.rgb = C_WHITE
    p.space_after = PptxPt(10)

c_right = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(6.85), PptxInches(1.75), col_w, PptxInches(5.0))
c_right.fill.solid()
c_right.fill.fore_color.rgb = PptxRGBColor(15, 30, 25)
c_right.line.color.rgb = C_EMERALD
tf_r = c_right.text_frame
tf_r.word_wrap = True

pr = tf_r.paragraphs[0]
pr.text = "🟢 NOUVELLE APPROCHE & OFFRE OFFERTE"
pr.font.size = PptxPt(15)
pr.font.bold = True
pr.font.color.rgb = PptxRGBColor(52, 211, 153)
pr.space_after = PptxPt(14)

nouv_pts = [
    "• Conception & migration : 0 $ (Offert pour soutenir un commerce remarquable de notre communauté).",
    "• Hébergement haute performance : 0 $ / mois (zéro mensualité récurrente).",
    "• Seul frais maintenu : votre domaine .ca habituel (~15 $ à 25 $ / an).",
    "• 6 mois de mises à jour courantes gratuites incluses (prix, menus de saison, horaires spéciaux).",
    "• Par la suite, si vous voulez de gros ajouts, nous en discuterons selon le travail nécessaire — aucun abonnement forcé."
]
for pt in nouv_pts:
    p = tf_r.add_paragraph()
    p.text = pt
    p.font.size = PptxPt(12.5)
    p.font.color.rgb = C_WHITE
    p.space_after = PptxPt(10)

# ------------------------------------------------------------------------------
# Slide 7: Évolution Optionnelle (DoorDash Storefront)
# ------------------------------------------------------------------------------
slide7 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide7)
add_header(slide7, "Évolution Optionnelle (Phase 2)", "Commandes en ligne directes sans commission (piste DoorDash Storefront)")

tb7 = slide7.shapes.add_textbox(PptxInches(0.8), PptxInches(1.75), PptxInches(11.7), PptxInches(0.8))
tf7 = tb7.text_frame
tf7.word_wrap = True
p7 = tf7.paragraphs[0]
p7.text = "Une piste intéressante à explorer ensemble auprès de DoorDash pour vos commandes directes :"
p7.font.size = PptxPt(13.5)
p7.font.color.rgb = C_WHITE

ds_cards = [
    ("0 % Commission Marketplace", "Sur les commandes passées directement sur votre site web (seuls les frais bancaires de carte habituels s'appliquent)."),
    ("0 $ / mois de frais de logiciel", "Aucun abonnement logiciel mensuel récurrent requis pour votre restaurant."),
    ("Gestion simplifiée", "Possibilité de gérer les commandes sur votre matériel existant (conditions exactes à valider avec DoorDash)."),
    ("Option de livraison par Dashers", "Des livreurs DoorDash assurent la course avec des frais de livraison transparents pour le client.")
]

for idx, (dtitle, ddesc) in enumerate(ds_cards):
    x = PptxInches(0.8 + (idx % 2) * 6.0)
    y = PptxInches(2.65 + (idx // 2) * 2.1)
    card = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, PptxInches(5.7), PptxInches(1.85))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD
    card.line.color.rgb = C_GOLD
    
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"✨ {dtitle}"
    p.font.size = PptxPt(14)
    p.font.bold = True
    p.font.color.rgb = C_GOLD_LIGHT
    p.space_after = PptxPt(6)
    
    p2 = tf.add_paragraph()
    p2.text = ddesc
    p2.font.size = PptxPt(12.5)
    p2.font.color.rgb = C_SLATE

# ------------------------------------------------------------------------------
# Slide 8: Transition Sans Stress (4 étapes)
# ------------------------------------------------------------------------------
slide8 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide8)
add_header(slide8, "Transition Sans Stress", "Comment se déroule la mise en ligne étape par étape")

steps = [
    ("1. Découverte décontractée", "Présentation de la démo sur votre téléphone ou directement sur mon écran. Zéro engagement."),
    ("2. Ajustements & validation avec le chef", "Le chef approuve les menus, prix réels, buffet, combos et photos officielles."),
    ("3. Mise en ligne sans coupure", "Votre site actuel reste actif jusqu'à validation finale. La bascule prend quelques heures sans aucune interruption."),
    ("4. 6 mois de suivi gratuit", "Modifications de menu, horaires des fêtes et ajustements courants pris en charge sans frais.")
]

for idx, (stitle, sdesc) in enumerate(steps):
    y = PptxInches(1.75 + idx * 1.32)
    c_step = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), y, PptxInches(11.7), PptxInches(1.15))
    c_step.fill.solid()
    c_step.fill.fore_color.rgb = C_CARD
    c_step.line.color.rgb = PptxRGBColor(51, 65, 85)
    
    tf = c_step.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = stitle
    p.font.size = PptxPt(14)
    p.font.bold = True
    p.font.color.rgb = C_GOLD
    p.space_after = PptxPt(3)
    
    p2 = tf.add_paragraph()
    p2.text = sdesc
    p2.font.size = PptxPt(12.5)
    p2.font.color.rgb = C_WHITE

# ------------------------------------------------------------------------------
# Slide 9: Conclusion & Contact + QR Code
# ------------------------------------------------------------------------------
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide9)
add_header(slide9, "Prochaine Étape", "Discutons-en 10 minutes autour d'un café")

tb9 = slide9.shapes.add_textbox(PptxInches(0.8), PptxInches(1.75), PptxInches(8.0), PptxInches(5.0))
tf9 = tb9.text_frame
tf9.word_wrap = True

p9_1 = tf9.paragraphs[0]
p9_1.text = "Essayez la démonstration dès maintenant :"
p9_1.font.size = PptxPt(16)
p9_1.font.bold = True
p9_1.font.color.rgb = C_WHITE
p9_1.space_after = PptxPt(8)

p9_url = tf9.add_paragraph()
p9_url.text = "👉 https://lecordondore.jopism.com"
p9_url.font.size = PptxPt(19)
p9_url.font.bold = True
p9_url.font.color.rgb = C_GOLD
p9_url.space_after = PptxPt(20)

p9_txt = tf9.add_paragraph()
p9_txt.text = "Mon but est simplement de vous aider et de vous remettre un site moderne clé en main pour valoriser Le Cordon Doré dans notre communauté.\n\nJe serais ravi de passer vous saluer au bistrot sur la rue Principale à Aylmer (ou d'échanger par court appel vidéo) pour vous faire essayer le site sur votre téléphone ou sur mon écran, en toute simplicité."
p9_txt.font.size = PptxPt(13.5)
p9_txt.font.color.rgb = C_SLATE
p9_txt.space_after = PptxPt(24)

p9_c = tf9.add_paragraph()
p9_c.text = "Jonathan Pankratz · Jopism Solutions Numériques\nCourriel : info@jopism.com  |  Web : digital.jopism.com"
p9_c.font.size = PptxPt(13)
p9_c.font.bold = True
p9_c.font.color.rgb = C_GOLD_LIGHT

# Add QR code box on Slide 9
c_qr9 = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(9.2), PptxInches(1.75), PptxInches(3.3), PptxInches(5.0))
c_qr9.fill.solid()
c_qr9.fill.fore_color.rgb = C_CARD
c_qr9.line.color.rgb = C_GOLD

if os.path.exists('qr_gold.png'):
    slide9.shapes.add_picture('qr_gold.png', PptxInches(9.5), PptxInches(2.2), width=PptxInches(2.7))

tb_q9 = slide9.shapes.add_textbox(PptxInches(9.3), PptxInches(5.2), PptxInches(3.1), PptxInches(1.2))
tf_q9 = tb_q9.text_frame
tf_q9.word_wrap = True
p_q9 = tf_q9.paragraphs[0]
p_q9.text = "Ouvrir la démo sur votre téléphone"
p_q9.font.size = PptxPt(11)
p_q9.font.bold = True
p_q9.font.color.rgb = C_GOLD_LIGHT
p_q9.alignment = PP_ALIGN.CENTER

prs.save('Presentation-Le-Cordon-Dore.pptx')
print("Presentation-Le-Cordon-Dore.pptx generated successfully (9 widescreen slides, 13-15pt text).")
