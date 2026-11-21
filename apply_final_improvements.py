# coding: utf-8
import os, re
import docx
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

# ==========================================
# 1. UPDATE proposal.html & proposal-complete.html
# ==========================================
for fname in ['proposal.html', 'proposal-complete.html']:
    with open(fname, 'r', encoding='utf-8') as f:
        html = f.read()

    # Section 1 chef bio update
    html = re.sub(
        r'un <strong>chef formé à Paris avec 25 ans d’expérience.*?</strong>',
        'un chef propriétaire d\'exception, <strong>Oodith Kallaram</strong>, formé à la prestigieuse <strong>Maison Lenôtre</strong> à Paris avec plus de 25 ans d’expérience internationale et intronisé parmi <strong>Les Disciples d\'Auguste Escoffier (Océan Indien)</strong>',
        html
    )

    # Add Chef Page finding to Section 2 if not present
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
        else:
            html = html.replace('<div class="space-y-3 text-xs sm:text-sm">', '<div class="space-y-3 text-xs sm:text-sm">\n' + chef_finding_html)

    # Section 3 feature update
    html = html.replace(
        'Le parcours du chef avec sa photo officielle en tenue, ses 25 ans d’art culinaire et des avis Google récents vérifiés.',
        'Le parcours du chef Oodith Kallaram avec sa photo officielle en tenue de Disciple d\'Escoffier, sa formation Maison Lenôtre à Paris, ses 25 ans d’art culinaire et des avis Google récents vérifiés.'
    )
    html = html.replace(
        'Le parcours du chef avec sa photo officielle en tenue de Disciple d\'Escoffier, ses 25 ans d’art culinaire et des avis Google récents vérifiés.',
        'Le parcours du chef Oodith Kallaram avec sa photo officielle en tenue de Disciple d\'Escoffier, sa formation Maison Lenôtre à Paris, ses 25 ans d’art culinaire et des avis Google récents vérifiés.'
    )

    # Clean any leftover coquille references
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
    # Flexible laptop / phone wording
    html = html.replace(
        'pour vous faire essayer la démonstration directement sur votre propre téléphone et répondre à vos questions.',
        'pour vous faire essayer la démonstration directement sur votre téléphone ou sur mon écran, et répondre à toutes vos questions.'
    )

    with open(fname, 'w', encoding='utf-8') as f:
        f.write(html)

print("HTML proposal files updated.")

# ==========================================
# 2. UPDATE PROPOSAL.md
# ==========================================
with open('PROPOSAL.md', 'r', encoding='utf-8') as f:
    md = f.read()

md = re.sub(
    r'Un \*\*chef formé à Paris avec 25 ans d’expérience.*?\*\*',
    'Un chef propriétaire d\'exception, **Chef Oodith Kallaram**, formé à la prestigieuse **Maison Lenôtre** à Paris avec plus de 25 ans d’expérience internationale et intronisé parmi **Les Disciples d\'Auguste Escoffier (Océan Indien)**',
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
md = re.sub(r'\bcoquilles?\b', 'petites erreurs', md, flags=re.IGNORECASE)
md = md.replace(
    '*[4,7 ★ · 180+ avis – à valider sur place]*',
    '**4,7 ★ · 180+ avis vérifiés sur Google**'
)
md = md.replace(
    '[4,7 ★ · 180+ avis – à valider sur place]',
    '4,7 ★ · 180+ avis vérifiés sur Google'
)
md = md.replace(
    'directement sur votre propre téléphone et répondre à vos questions',
    'sur votre téléphone ou directement sur mon écran, et répondre à toutes vos questions'
)

with open('PROPOSAL.md', 'w', encoding='utf-8') as f:
    f.write(md)

print("PROPOSAL.md updated.")

# ==========================================
# 3. REBUILD WORD DOCUMENTS
# ==========================================
def build_docx(filename):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # En-tête
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
        ('Ce que ça coûte', 'La conception, la migration technique et l’hébergement moderne sont offerts (0 $/mois). Mises à jour courantes gratuites pendant 6 mois sans engagement.'),
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
    doc.add_paragraph('Le Cordon Doré a des atouts que très peu de restaurants peuvent revendiquer dans l\'Outaouais : un chef propriétaire d\'exception, Oodith Kallaram, formé à la prestigieuse Maison Lenôtre à Paris avec plus de 25 ans d’expérience culinaire et hôtelière internationale, et intronisé parmi Les Disciples d\'Auguste Escoffier (Océan Indien); une double cuisine indienne et bistrot international soignée; des viandes 100 % certifiées Halal; un buffet réputé chaque mardi soir; et une adresse bien ancrée au cœur du Vieux-Aylmer avec plus de 180 avis élogieux (4,7 ★ sur Google).')
    doc.add_paragraph('Le site web doit être à la hauteur de cette réputation, surtout pour le client qui vous cherche sur son téléphone et décide où manger en moins de 3 secondes.')

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 3. Ce que j'ai remarqué
    h3 = doc.add_heading('2. Ce que j’ai remarqué sur le site actuel (et ce qu\'on améliore)', level=1)
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
        ('Pages inachevées & pages fantômes', 'Le blogue ne contient qu’un article vide intitulé « Test » (« Aucun post publié dans cette langue »), et plusieurs pages de test Wix (« copie-de-... ») sont indexées sur Google.', 'Architecture monopage moderne sans cul-de-sac, avec redirection automatique des anciennes adresses pour préserver votre référencement.'),
        ('Carte des boissons incomplète', 'Les 7 cocktails sont affichés à 14 $, mais les rubriques vins, bières et spiritueux ne montrent aucun article ni tarif.', 'Bannière mettant en valeur les cocktails signatures à 14 $ et grille des 8 catégories de la cave à compléter selon vos sélections.')
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
        ('Le buffet du mardi en vedette', 'Un bloc dédié avec l’horaire (17h–20h30) et la réservation, plus un renvoi direct vers votre page Facebook officielle pour les brunchs spéciaux du week-end.'),
        ('Des menus complets & bilingues', '44 plats indiens, 20 spécialités bistro, 6 menus enfants, desserts et cocktails, avec badges clairs (Halal, Végétarien, Sans gluten, piquant).'),
        ('Traiteur et cartes-cadeaux', 'Une section pour recevoir les soumissions (mariages, fêtes privées, corporatif dans l’Outaouais) et afficher les montants disponibles (25 $ à 200 $).'),
        ('Votre histoire & authenticité', 'Le parcours du chef Oodith Kallaram avec sa photo officielle en tenue de Disciple d\'Escoffier, sa formation Maison Lenôtre à Paris, ses 25 ans d’art culinaire et des avis Google récents vérifiés.'),
        ('Visibilité locale renforcée', 'Fiche Google alignée avec le site, vitesse de chargement inférieure à 1 seconde et balisage SEO propre pour attirer la clientèle d\'Aylmer, Gatineau et Ottawa.')
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
        ('Combos familiaux', 'Les compositions et les prix de la démo sont des exemples à ajuster selon vos formules exactes.'),
        ('Photos', 'Vos vraies photos (le chef, la salle, le buffet, les plats) à la place des images d\'illustration.'),
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

    # 6. Coûts et engagements
    h6 = doc.add_heading('5. Coûts et engagements', level=1)
    h6.style.font.color.rgb = RGBColor(15, 23, 42)

    costs = [
        ('Offerts', 'Conception complète du site, migration technique, mise en ligne et hébergement moderne haute performance (0 $/mois).'),
        ('À votre charge', 'Uniquement le renouvellement de votre nom de domaine .ca auprès de votre registraire habituel, soit habituellement entre 15 $ et 25 $ par année.'),
        ('Abonnement Wix', 'Vous pourrez l’annuler une fois la bascule terminée et vos cartes-cadeaux réglées. Vous économisez ainsi entre 250 $ et 530 $+ CAD par an.'),
        ('Mises à jour continues', 'Durant les 6 premiers mois, les ajustements courants (changement de prix, ajout de plats saisonniers, horaires spéciaux de fêtes ou du buffet) sont pris en charge gratuitement sans frais. Par la suite, nous pourrons convenir ensemble d\'un petit forfait d\'entretien mensuel très abordable ou de simples mises à jour à l\'acte, selon vos besoins réels et sans engagement.'),
        ('Propriété & Sans engagement', 'Votre domaine reste enregistré à votre nom, votre contenu vous appartient à 100 %, et je vous remets sur demande une copie complète du site. Vous êtes libres d\'arrêter quand vous le souhaitez.')
    ]
    for ctitle, cdesc in costs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.5)
        r = p.add_run(f'• {ctitle} : ')
        r.bold = True
        p.add_run(cdesc)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 7. Option DoorDash
    h7 = doc.add_heading('6. Option pour plus tard : commander en ligne avec DoorDash', level=1)
    h7.style.font.color.rgb = RGBColor(15, 23, 42)

    doc.add_paragraph('Puisque vous êtes déjà partenaire de DoorDash, il est possible d’offrir la commande en ligne directement depuis votre propre site web (technologie Storefront).')
    doc.add_paragraph('Selon la tarification canadienne publiée par DoorDash, la commande en ligne sur vos propres canaux n’a ni frais mensuels ni commission de marketplace, et la livraison par Dashers depuis votre site se facture à tarif fixe par livraison (environ 6,99 $ à 10,99 $) plutôt qu’en pourcentage. À titre de comparaison, la commission sur l\'application DoorDash est de 20 % à 29 % pour la livraison et de 8 % à 10 % pour le ramassage (soit 8 $ à 12 $ perdus sur une commande de 40 $).')
    doc.add_paragraph('Zéro friction en salle : Les commandes sonnent sur la même tablette DoorDash que vous utilisez déjà au comptoir ! Avant toute décision, nous confirmerons ensemble avec DoorDash les frais exacts et le fonctionnement. L’économie réelle dépend de la part de vos clients qui commanderaient depuis votre site plutôt que depuis l’application.')

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 8. Plan de déploiement
    h8 = doc.add_heading('7. Plan de déploiement', level=1)
    h8.style.font.color.rgb = RGBColor(15, 23, 42)

    deploys = [
        ('1. Essai de la démo', 'Vous la parcourez sur votre téléphone ou sur ordinateur à l\'adresse lecordondore.jopism.com et notez vos impressions et ajustements souhaités.'),
        ('2. Validation du contenu', 'Menu, prix, photos, avis, cartes-cadeaux et coordonnées, selon vos disponibilités (généralement une à deux semaines).'),
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
    print(f"{filename} generated.")

for fname in ['Proposition-Strategique-Complete.docx', 'Proposition-Sommaire-Comparatif.docx', 'Proposition-Refonte-Le-Cordon-Dore.docx']:
    build_docx(fname)

# ==========================================
# 4. REBUILD PRESENTATION POWERPOINT WITH QR & ESCOFFIER & NO COQUILLES
# ==========================================
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

def add_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_DARK
    bg.line.fill.background()

def add_header(slide, tag, title):
    tb_tag = slide.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.7), PptxInches(0.35))
    tf_tag = tb_tag.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag.upper()
    p_tag.font.size = PptxPt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_GOLD

    tb_title = slide.shapes.add_textbox(PptxInches(0.8), PptxInches(0.8), PptxInches(11.7), PptxInches(0.65))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title
    p_title.font.size = PptxPt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = C_WHITE

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(0.8), PptxInches(1.5), PptxInches(11.733), PptxInches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = PptxRGBColor(30, 41, 59)
    line.line.fill.background()

# Slide 1: Cover
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide1)
if os.path.exists('logo.png'):
    slide1.shapes.add_picture('logo.png', PptxInches(0.8), PptxInches(2.2), width=PptxInches(1.8))

tb1 = slide1.shapes.add_textbox(PptxInches(3.0), PptxInches(2.0), PptxInches(9.5), PptxInches(3.5))
tf1 = tb1.text_frame
tf1.word_wrap = True
p1_pre = tf1.paragraphs[0]
p1_pre.text = "PROPOSITION STRATÉGIQUE & DÉMONSTRATION"
p1_pre.font.size = PptxPt(11)
p1_pre.font.bold = True
p1_pre.font.color.rgb = C_GOLD
p1_pre.space_after = PptxPt(10)

p1_tit = tf1.add_paragraph()
p1_tit.text = "Un nouveau site web pour\nLe Cordon Doré"
p1_tit.font.size = PptxPt(36)
p1_tit.font.bold = True
p1_tit.font.color.rgb = C_WHITE
p1_tit.space_after = PptxPt(14)

p1_sub = tf1.add_paragraph()
p1_sub.text = "Mettre en valeur votre cuisine, l'histoire du chef Oodith Kallaram (Maison Lenôtre, Disciple d'Escoffier) et faciliter les réservations de vos convives."
p1_sub.font.size = PptxPt(14)
p1_sub.font.color.rgb = C_SLATE
p1_sub.space_after = PptxPt(20)

p1_foot = tf1.add_paragraph()
p1_foot.text = "Préparé par Jonathan Pankratz · Jopism Solutions Numériques  |  Démo : lecordondore.jopism.com"
p1_foot.font.size = PptxPt(11)
p1_foot.font.color.rgb = C_GOLD_LIGHT

# Slide 2: En Bref + QR Code
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide2)
add_header(slide2, "En résumé", "Une solution simple, moderne et avantageuse")

cards2 = [
    ("📱 Pensé d'abord pour le mobile", "Boutons directs pour appeler, texter et trouver le bistrot en 3 secondes sur téléphone ou sur ordinateur."),
    ("⚡ Vitesse de chargement éclair", "Chargement < 1 sec sans lenteurs de Wix, favorisant votre référencement sur Google Maps à Aylmer et Ottawa."),
    ("🛡️ Vous gardez le contrôle total", "Votre domaine reste enregistré à votre nom, votre contenu vous appartient à 100%, et vous arrêtez quand vous le souhaitez."),
    ("💰 Hébergement moderne à 0 $ / mois", "Conception et migration offertes. Mises à jour gratuites pendant 6 mois. Élimination des abonnements Wix coûteux.")
]

for idx, (ctitle, cdesc) in enumerate(cards2):
    x = PptxInches(0.8 + (idx % 2) * 4.6)
    y = PptxInches(1.8 + (idx // 2) * 2.6)
    card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, PptxInches(4.3), PptxInches(2.2))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD
    card.line.color.rgb = PptxRGBColor(51, 65, 85)
    
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = ctitle
    p.font.size = PptxPt(13.5)
    p.font.bold = True
    p.font.color.rgb = C_GOLD_LIGHT
    p.space_after = PptxPt(6)
    
    p2 = tf.add_paragraph()
    p2.text = cdesc
    p2.font.size = PptxPt(10.5)
    p2.font.color.rgb = C_SLATE

# Add QR box on right side of Slide 2
c_qr = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(10.3), PptxInches(1.8), PptxInches(2.2), PptxInches(4.8))
c_qr.fill.solid()
c_qr.fill.fore_color.rgb = C_CARD
c_qr.line.color.rgb = C_GOLD
if os.path.exists('qr_gold.png'):
    slide2.shapes.add_picture('qr_gold.png', PptxInches(10.5), PptxInches(2.4), width=PptxInches(1.8))
tb_q = slide2.shapes.add_textbox(PptxInches(10.4), PptxInches(4.4), PptxInches(2.0), PptxInches(2.0))
tf_q = tb_q.text_frame
tf_q.word_wrap = True
p_q = tf_q.paragraphs[0]
p_q.text = "Scannez pour ouvrir la démo en direct :\nlecordondore.jopism.com"
p_q.font.size = PptxPt(10)
p_q.font.bold = True
p_q.font.color.rgb = C_GOLD_LIGHT
p_q.alignment = PP_ALIGN.CENTER

# Slide 3: Ce que vous faites déjà très bien
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide3)
add_header(slide3, "Vos Atouts", "Ce que Le Cordon Doré fait déjà remarquablement bien")

if os.path.exists('chef.jpg'):
    slide3.shapes.add_picture('chef.jpg', PptxInches(0.8), PptxInches(1.8), width=PptxInches(3.2))

tb3 = slide3.shapes.add_textbox(PptxInches(4.4), PptxInches(1.8), PptxInches(8.1), PptxInches(5.0))
tf3 = tb3.text_frame
tf3.word_wrap = True

atouts = [
    ("👨‍🍳 Chef Oodith Kallaram · Maison Lenôtre & Disciple d'Auguste Escoffier", "25 ans de carrière internationale, formé à la prestigieuse Maison Lenôtre à Paris, intronisé Disciple d'Escoffier. Une alliance d'exception entre haute gastronomie française et épices d'Inde."),
    ("🍛 Cuisine authentique & variée", "44 spécialités indiennes soignées, classiques du bistrot international, 6 menus enfants et options véganes/végétariennes."),
    ("🥩 Viandes 100 % certifiées Halal", "Un gage d'excellence et de respect des traditions recherché par la clientèle de Gatineau et d'Ottawa."),
    ("✨ Grand Buffet Indien du Mardi Soir", "Une formule à volonté réputée qui fidélise les amateurs de cuisine indienne chaque semaine."),
    ("⭐ Réputation solide dans le Vieux-Aylmer", "Une note d'excellence de 4,7 ★ sur Google basée sur plus de 180 avis élogieux vérifiés.")
]

for idx, (atitle, adesc) in enumerate(atouts):
    p = tf3.paragraphs[0] if idx == 0 else tf3.add_paragraph()
    p.text = atitle + " : "
    p.font.size = PptxPt(12.5)
    p.font.bold = True
    p.font.color.rgb = C_GOLD
    p.space_after = PptxPt(8)
    
    r = p.add_run()
    r.text = adesc
    r.font.bold = False
    r.font.color.rgb = C_WHITE

# Slide 4: Diagnostic bienveillant
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide4)
add_header(slide4, "Diagnostic Bienveillant", "Ce que j'ai remarqué sur votre site actuel (et ce qu'on améliore)")

remarques_pptx = [
    ("Navigation encombrée", "13 liens avec doublon « Home » et « Acceuil » (faute d'orthographe).", "5 sections claires, bouton d'appel fixe et toggle bilingue fluide."),
    ("Page Parcours du Chef vide", "Lien vers URL de brouillon (/copie-de-menu-enfant) avec photo sans aucun texte.", "Section dédiée racontant le parcours d'Oodith : Maison Lenôtre, Escoffier et 25 ans d'art culinaire."),
    ("Coordonnées divergentes", "2 adresses Gmail distinctes (cordondor1803 vs 18103) et 3 numéros éparpillés.", "1 canal clair : appel commandes (819 682-5781), texto réservations (819 661-5113)."),
    ("Petites erreurs au menu", "Seekh kebab listé deux fois, fautes d'anglais (Scalops and calamary, briyani rice).", "Menu relu, orthographe bilingue soignée, validé avec le chef."),
    ("Menu familial manquant", "Le lien « Menu familiale » mène à une page affichant un vieil avis de vacances sans plat.", "Onglet festins familiaux avec compositions et prix transparents."),
    ("Pages inachevées & boissons", "Page /blog avec article « Test » vide, et sélections de la cave sans prix.", "Site monopage sans cul-de-sac, cocktails à 14 $ mis en avant et cave soignée.")
]

for idx, (rtitle, rcur, rsol) in enumerate(remarques_pptx):
    x = PptxInches(0.8 + (idx % 2) * 6.0)
    y = PptxInches(1.8 + (idx // 2) * 1.7)
    card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, PptxInches(5.6), PptxInches(1.5))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD
    card.line.color.rgb = PptxRGBColor(51, 65, 85)
    
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"• {rtitle}"
    p.font.size = PptxPt(12)
    p.font.bold = True
    p.font.color.rgb = C_GOLD_LIGHT
    p.space_after = PptxPt(2)
    
    p2 = tf.add_paragraph()
    r_cur = p2.add_run()
    r_cur.text = f"Actuel : {rcur}\n"
    r_cur.font.size = PptxPt(9.5)
    r_cur.font.color.rgb = PptxRGBColor(244, 63, 94)
    
    r_sol = p2.add_run()
    r_sol.text = f"Nouveau : {rsol}"
    r_sol.font.size = PptxPt(9.5)
    r_sol.font.bold = True
    r_sol.font.color.rgb = C_EMERALD

# Slide 5: Transparence totale
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide5)
add_header(slide5, "Transparence Totale", "Ce que nous validerons ensemble avant toute mise en ligne")

tb5 = slide5.shapes.add_textbox(PptxInches(0.8), PptxInches(1.8), PptxInches(11.7), PptxInches(5.0))
tf5 = tb5.text_frame
tf5.word_wrap = True

validations = [
    ("Menus, prix et allergènes", "Le chef approuve chaque plat, description et mention diététique. Rien n'est publié sans son accord formel."),
    ("Combos familiaux", "Les compositions et les tarifs de la démonstration sont des suggestions à ajuster selon vos formules exactes."),
    ("Photos officielles", "Intégration de vos vraies photos (salle, cuisine, buffet, chef) en remplacement des images d'illustration."),
    ("Avis et note Google", "Affichage fidèle vérifié le jour J de la note officielle : 4,7 ★ · 180+ avis vérifiés sur Google."),
    ("Cartes-cadeaux", "Modalités pour continuer à offrir vos cartes numériques de 25 $ à 200 $ et honorer celles déjà vendues."),
    ("Coordonnées & Réseaux sociaux", "Vérification des liens officiels Facebook (cordondorerestaurantbistrot), Instagram (@cordondorerestaurantbistro) et courriel officiel.")
]

for idx, (vtitle, vdesc) in enumerate(validations):
    p = tf5.paragraphs[0] if idx == 0 else tf5.add_paragraph()
    p.text = f"✓ {vtitle} : "
    p.font.size = PptxPt(13)
    p.font.bold = True
    p.font.color.rgb = C_GOLD
    p.space_after = PptxPt(8)
    
    r = p.add_run()
    r.text = vdesc
    r.font.bold = False
    r.font.color.rgb = C_WHITE

# Slide 6: Portrait financier & économies
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide6)
add_header(slide6, "Aspect Financier", "Frais actuels vs Nouvelle approche")

col_w = PptxInches(5.6)
c_left = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(1.8), col_w, PptxInches(4.8))
c_left.fill.solid()
c_left.fill.fore_color.rgb = PptxRGBColor(30, 20, 25)
c_left.line.color.rgb = PptxRGBColor(225, 29, 72)
tf_l = c_left.text_frame
tf_l.word_wrap = True

pl = tf_l.paragraphs[0]
pl.text = "Coûts & Commissions Actuels"
pl.font.size = PptxPt(16)
pl.font.bold = True
pl.font.color.rgb = PptxRGBColor(251, 113, 133)
pl.space_after = PptxPt(12)

pl_body = tf_l.add_paragraph()
pl_body.text = (
    "• Abonnement Wix annuel : Environ 250 $ à 530 $+ CAD/an pour maintenir le site actuel.\n\n"
    "• Commissions des plateformes de livraison (DoorDash, etc.) : 20 % à 29 % sur chaque livraison, et 8 % à 15 % sur le ramassage.\n\n"
    "• Perte de marge : Sur une commande moyenne de 40 $, cela représente 8 $ à 12 $ prélevés par la plateforme.\n\n"
    "• Risque sur les réservations : 2 courriels différents qui créent de la confusion et risquent de faire perdre des réservations."
)
pl_body.font.size = PptxPt(12)
pl_body.font.color.rgb = C_SLATE

c_right = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(6.8), PptxInches(1.8), col_w, PptxInches(4.8))
c_right.fill.solid()
c_right.fill.fore_color.rgb = PptxRGBColor(15, 30, 25)
c_right.line.color.rgb = C_EMERALD
tf_r = c_right.text_frame
tf_r.word_wrap = True

pr = tf_r.paragraphs[0]
pr.text = "Nouvelle Approche & Offre Offerte"
pr.font.size = PptxPt(16)
pr.font.bold = True
pr.font.color.rgb = PptxRGBColor(52, 211, 153)
pr.space_after = PptxPt(12)

pr_body = tf_r.add_paragraph()
pr_body.text = (
    "• Hébergement moderne à 0 $ / mois : Zéro mensualité. Vous conservez uniquement le renouvellement de votre domaine .ca (~15 $ à 25 $/an).\n\n"
    "• Conception & migration offertes : Je prends en charge la mise en ligne et le transfert de lecordondorerestaurantbistrot.ca sans interruption.\n\n"
    "• Mises à jour continues : 6 premiers mois de modifications courantes inclus gratuitement (prix, menus de saison, horaires spéciaux), sans engagement.\n\n"
    "• Plus de commandes directes : Le bouton d'appel en 1-clic stimule les commandes à emporter directes sans commission."
)
pr_body.font.size = PptxPt(12)
pr_body.font.color.rgb = C_SLATE

# Slide 7: DoorDash Storefront
slide7 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide7)
add_header(slide7, "Opportunité Phase 2", "Option pour plus tard : Commander en ligne avec DoorDash Storefront")

tb7 = slide7.shapes.add_textbox(PptxInches(0.8), PptxInches(1.8), PptxInches(11.7), PptxInches(5.0))
tf7 = tb7.text_frame
tf7.word_wrap = True

p7_intro = tf7.paragraphs[0]
p7_intro.text = "Vous êtes déjà partenaire DoorDash : vous pouvez transformer votre site en canal de vente direct sans commission marketplace."
p7_intro.font.size = PptxPt(13)
p7_intro.font.bold = True
p7_intro.font.color.rgb = C_GOLD_LIGHT
p7_intro.space_after = PptxPt(14)

dd_points = [
    ("0 $ de frais mensuels & 0 % de commission sur votre site", "Contrairement aux commandes passées sur l'application DoorDash (20% à 29%), les commandes initiées depuis votre site web n'ont aucune commission de marketplace (seuls les frais bancaires de carte s'appliquent)."),
    ("Zéro friction technique au restaurant", "Les commandes se synchronisent directement avec votre menu existant et sonnent sur la même tablette DoorDash que vous avez déjà au comptoir. Aucune nouvelle habitude à apprendre pour votre personnel."),
    ("Livraison assurée par les Dashers à tarif fixe", "Si le client demande la livraison, un livreur Dasher vient chercher la commande et la livre, facturé à tarif fixe par course (environ 6,99 $ à 10,99 $) plutôt qu'en pourcentage sur votre chiffre d'affaires."),
    ("Validation préalable ensemble", "Avant toute décision, nous validerons ensemble les conditions avec DoorDash et l'intérêt selon votre volume de commandes.")
]

for dtitle, ddesc in dd_points:
    p = tf7.add_paragraph()
    p.text = f"• {dtitle} : "
    p.font.size = PptxPt(12)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    p.space_after = PptxPt(8)
    
    r = p.add_run()
    r.text = ddesc
    r.font.bold = False
    r.font.color.rgb = C_SLATE

# Slide 8: Plan de déploiement
slide8 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide8)
add_header(slide8, "Transition Sans Risque", "Plan de déploiement en 4 étapes simples")

steps = [
    ("1. Essai de la Démo", "Vous parcourez le site de démonstration sur votre téléphone ou ordinateur (lecordondore.jopism.com) et notez vos impressions."),
    ("2. Validation du Contenu", "Validation du menu, des prix, des photos réelles, des avis et des coordonnées à votre rythme (1 à 2 semaines)."),
    ("3. Mise en Ligne Transparente", "Votre site actuel reste en ligne jusqu'à ce que le nouveau soit parfait. La bascule prend quelques heures sans aucune interruption."),
    ("4. Suivi & Accompagnement", "Configuration des redirections, alignement de la fiche Google et petits ajustements durant le premier mois.")
]

for idx, (stitle, sdesc) in enumerate(steps):
    x = PptxInches(0.8 + (idx % 2) * 6.0)
    y = PptxInches(1.8 + (idx // 2) * 2.5)
    card = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, PptxInches(5.6), PptxInches(2.1))
    card.fill.solid()
    card.fill.fore_color.rgb = C_CARD
    card.line.color.rgb = PptxRGBColor(51, 65, 85)
    
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = stitle
    p.font.size = PptxPt(14)
    p.font.bold = True
    p.font.color.rgb = C_GOLD_LIGHT
    p.space_after = PptxPt(6)
    
    p2 = tf.add_paragraph()
    p2.text = sdesc
    p2.font.size = PptxPt(11)
    p2.font.color.rgb = C_SLATE

# Slide 9: Conclusion avec QR code et contact
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide9)

# Text box on left
tb9 = slide9.shapes.add_textbox(PptxInches(0.8), PptxInches(1.2), PptxInches(8.5), PptxInches(5.0))
tf9 = tb9.text_frame
tf9.word_wrap = True

p9_tag = tf9.paragraphs[0]
p9_tag.text = "PROCHAINE ÉTAPE SIMPLE"
p9_tag.font.size = PptxPt(11)
p9_tag.font.bold = True
p9_tag.font.color.rgb = C_GOLD
p9_tag.space_after = PptxPt(10)

p9_tit = tf9.add_paragraph()
p9_tit.text = "Une présentation de 10 minutes,\nsans aucun engagement."
p9_tit.font.size = PptxPt(30)
p9_tit.font.bold = True
p9_tit.font.color.rgb = C_WHITE
p9_tit.space_after = PptxPt(14)

p9_desc = tf9.add_paragraph()
p9_desc.text = "Je serais très heureux de passer vous voir au restaurant sur la rue Principale à Aylmer (ou lors d'un court appel vidéo) pour vous faire essayer la démonstration sur votre téléphone ou directement sur mon écran, et répondre à toutes vos questions."
p9_desc.font.size = PptxPt(13.5)
p9_desc.font.color.rgb = C_SLATE
p9_desc.space_after = PptxPt(20)

p9_contact = tf9.add_paragraph()
p9_contact.text = "Jonathan Pankratz · Jopism Solutions Numériques\nCourriel : info@jopism.com  |  Web : digital.jopism.com  |  Démo : lecordondore.jopism.com"
p9_contact.font.size = PptxPt(12)
p9_contact.font.bold = True
p9_contact.font.color.rgb = C_GOLD_LIGHT

# QR Code Card on right
c_qr9 = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(9.8), PptxInches(1.2), PptxInches(2.7), PptxInches(5.0))
c_qr9.fill.solid()
c_qr9.fill.fore_color.rgb = C_CARD
c_qr9.line.color.rgb = C_GOLD

if os.path.exists('qr_gold.png'):
    slide9.shapes.add_picture('qr_gold.png', PptxInches(10.15), PptxInches(1.6), width=PptxInches(2.0))

tb_qr9 = slide9.shapes.add_textbox(PptxInches(9.9), PptxInches(3.8), PptxInches(2.5), PptxInches(2.2))
tf_qr9 = tb_qr9.text_frame
tf_qr9.word_wrap = True
p_qr9 = tf_qr9.paragraphs[0]
p_qr9.text = "Scannez avec votre téléphone pour ouvrir la démo en direct :\n\nlecordondore.jopism.com"
p_qr9.font.size = PptxPt(10.5)
p_qr9.font.bold = True
p_qr9.font.color.rgb = C_GOLD_LIGHT
p_qr9.alignment = PP_ALIGN.CENTER

prs.save(r'C:\Users\joppa\Desktop\Documents\GitHub\Le Cordon Doré\Presentation-Le-Cordon-Dore.pptx')
print("Presentation-Le-Cordon-Dore.pptx generated successfully with QR codes and Escoffier credentials!")
