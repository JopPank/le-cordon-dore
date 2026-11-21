import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_contact_block(doc):
    p_box = doc.add_paragraph()
    p_box.paragraph_format.space_before = Pt(12)
    p_box.paragraph_format.space_after = Pt(2)
    
    r_name = p_box.add_run('Jonathan Pankratz')
    r_name.bold = True
    r_name.font.size = Pt(10)
    r_name.font.color.rgb = RGBColor(15, 23, 42)
    
    r_sub = p_box.add_run(' · Jopism Solutions Numériques\n')
    r_sub.font.size = Pt(9)
    r_sub.font.color.rgb = RGBColor(71, 85, 105)
    
    r_c1 = p_box.add_run('Courriel : info@jopism.com   |   Web : digital.jopism.com   |   Démo : lecordondore.jopism.com\n')
    r_c1.font.size = Pt(8.5)
    r_c1.font.bold = True
    r_c1.font.color.rgb = RGBColor(180, 83, 9)
    
    r_prop = p_box.add_run('Offre sans engagement : ')
    r_prop.bold = True
    r_prop.font.size = Pt(8.5)
    
    r_txt = p_box.add_run('Disponible pour une courte présentation de 10 minutes (en personne dans votre restaurant à Aylmer ou par appel vidéo) pour vous montrer le fonctionnement en direct sur votre téléphone et répondre à vos questions.')
    r_txt.font.size = Pt(8.5)
    r_txt.font.color.rgb = RGBColor(51, 65, 85)

# ==========================================
# 1. GENERATE Proposition-Sommaire-Comparatif.docx
# ==========================================
doc_sum = docx.Document()
for s in doc_sum.sections:
    s.top_margin = Inches(0.45)
    s.bottom_margin = Inches(0.45)
    s.left_margin = Inches(0.5)
    s.right_margin = Inches(0.5)

p_sum_head = doc_sum.add_paragraph()
p_sum_head.paragraph_format.space_after = Pt(2)
r_sum_pre = p_sum_head.add_run('FICHE SYNTHÈSE & COMPARATIF EXÉCUTIF\n')
r_sum_pre.font.size = Pt(8.5)
r_sum_pre.font.bold = True
r_sum_pre.font.color.rgb = RGBColor(180, 83, 9)

r_sum_tit = p_sum_head.add_run('Le Cordon Doré Restaurant-Bistrot\n')
r_sum_tit.font.size = Pt(16)
r_sum_tit.font.bold = True
r_sum_tit.font.color.rgb = RGBColor(15, 23, 42)

r_sum_sub = p_sum_head.add_run('58, rue Principale, Gatineau (Vieux-Aylmer)  |  Démonstration en ligne : https://lecordondore.jopism.com')
r_sum_sub.font.size = Pt(9)

t_sum = doc_sum.add_table(rows=1, cols=3)
t_sum.alignment = WD_TABLE_ALIGNMENT.CENTER
t_sum.autofit = False

h_cells = t_sum.rows[0].cells
widths_sum = [Inches(1.8), Inches(2.7), Inches(2.9)]
headers_sum = ['Aspect Évalué', 'Site Actuel (Wix)', 'Nouvelle Version 2026']

for i, title in enumerate(headers_sum):
    h_cells[i].text = title
    h_cells[i].paragraphs[0].runs[0].font.bold = True
    h_cells[i].paragraphs[0].runs[0].font.size = Pt(8.5)
    set_cell_background(h_cells[i], 'F1F5F9')
    set_cell_margins(h_cells[i], 80, 80, 100, 100)
    h_cells[i].width = widths_sum[i]

summary_rows = [
    ('Navigation Mobile', '13 liens encombrés avec doublon et faute d\'orthographe ("Home" et "Acceuil" au lieu d\'"Accueil")', '5 destinations claires + bouton d\'appel direct visible en 1 clic'),
    ('Consultation des Menus', 'Menus dispersés; le lien "Menu familiale" mène à un avis de vacances sans aucun combo', 'Menus interactifs par onglets fluides avec prix et festins familiaux complets'),
    ('Exactitude du Menu', 'Plat en doublon (Seekh kebab poulet affiché 2 fois) et coquilles ("Scalops and calamary", "briyani rice")', 'Menu 100% vérifié, bilingue soigné, sans doublon ni faute d\'orthographe'),
    ('Pages Inachevées & Vides', 'Page /blog affichant un article "Test" sans texte et /cave-cordon listant des catégories sans prix', 'Site moderne sans cul-de-sac : carte des cocktails signature (14 $) et sélections claires'),
    ('Langues (Bilinguisme)', 'Mélange de textes bilingues sans toggle clair', 'Bouton FR / EN instantané pour toute la clientèle d\'Aylmer & Ottawa'),
    ('Prise de Commande & Contact', '3 numéros dispersés & faute dans le courriel ("cordondor1803" vs "cordondor18103")', 'Boutons dédiés : Appels directs (819 682-5781), Texto réservations (819 661-5113) et courriel vérifié'),
    ('Promotion Événements', 'Affiches images statiques souvent tronquées', 'Section vedette pour le Buffet Indien du mardi & service traiteur'),
    ('Confiance & Réputation', 'Les 4.8★ Google et la certification Halal sont peu visibles', 'Badges Halal, 4.8★ Google et histoire du Chef formé à Paris en tête de page')
]

for row in summary_rows:
    r_cells = t_sum.add_row().cells
    for i, text in enumerate(row):
        r_cells[i].text = text
        p = r_cells[i].paragraphs[0]
        p.runs[0].font.size = Pt(8)
        set_cell_margins(r_cells[i], 50, 50, 90, 90)
        r_cells[i].width = widths_sum[i]
        if i == 0:
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(15, 23, 42)
        elif i == 1:
            p.runs[0].font.color.rgb = RGBColor(190, 18, 60)
        elif i == 2:
            p.runs[0].font.color.rgb = RGBColor(4, 120, 87)
            p.runs[0].font.bold = True

doc_sum.add_paragraph().paragraph_format.space_after = Pt(4)

# COÛTS & TRANSITION SANS RISQUE Box
p_fin_box = doc_sum.add_paragraph()
p_fin_box.paragraph_format.space_after = Pt(2)
r_fb_tit = p_fin_box.add_run('PORTRAIT FINANCIER & TRANSITION SANS RISQUE (Hébergement : 0 $ / mois)\n')
r_fb_tit.bold = True
r_fb_tit.font.size = Pt(9)
r_fb_tit.font.color.rgb = RGBColor(180, 83, 9)

r_fb_c1 = p_fin_box.add_run('• Coûts actuels Wix & Livraison : ')
r_fb_c1.bold = True
p_fin_box.add_run('Abonnement Wix de ~250 $ à 530 $+ CAD/an. Commissions de 20% à 30% sur livraison et 8% à 15% sur ramassage via les applications tierces comme DoorDash (8 $ à 12 $ perdus sur une commande de 40 $).\n')

r_fb_c2 = p_fin_box.add_run('• Nouvelle approche économique : ')
r_fb_c2.bold = True
p_fin_box.add_run('Hébergement moderne à 0 $/mois. Vous conservez uniquement votre nom de domaine .ca (~15 $–25 $/an). Commandes directes stimulées par téléphone.\n')

r_fb_c3 = p_fin_box.add_run('• Mon offre gratuite & Phase 2 DoorDash : ')
r_fb_c3.bold = True
p_fin_box.add_run('Conception, migration et maintenance continue gratuites avec zéro interruption. Option Phase 2 : Commande en ligne DoorDash sur votre site à 0 $/mois et 0 % de commission marketplace, synchronisée sur votre tablette actuelle au comptoir avec livraison par Dashers sans les 20%–29% de commission.')

add_contact_block(doc_sum)
doc_sum.save(r'C:\Users\joppa\Desktop\Documents\GitHub\Le Cordon Doré\Proposition-Sommaire-Comparatif.docx')
print('Proposition-Sommaire-Comparatif.docx updated successfully')

# ==========================================
# 2. GENERATE Proposition-Strategique-Complete.docx
# ==========================================
doc_comp = docx.Document()
for s in doc_comp.sections:
    s.top_margin = Inches(0.55)
    s.bottom_margin = Inches(0.55)
    s.left_margin = Inches(0.6)
    s.right_margin = Inches(0.6)

p_c_sub = doc_comp.add_paragraph()
r_c_sub = p_c_sub.add_run('DOCUMENT DE PRÉSENTATION OFFICIELLE')
r_c_sub.font.size = Pt(8.5)
r_c_sub.font.bold = True
r_c_sub.font.color.rgb = RGBColor(180, 83, 9)

p_c_title = doc_comp.add_paragraph()
r_c_title = p_c_title.add_run('Proposition de Refonte de Site Web')
r_c_title.font.size = Pt(18)
r_c_title.font.bold = True
r_c_title.font.color.rgb = RGBColor(15, 23, 42)

p_c_rest = doc_comp.add_paragraph()
r_c_rest = p_c_rest.add_run('Restaurant-Bistrot Le Cordon Doré\n')
r_c_rest.font.size = Pt(12)
r_c_rest.font.bold = True
r_c_rest.font.color.rgb = RGBColor(180, 83, 9)

r_c_addr = p_c_rest.add_run('58, rue Principale, Gatineau (Vieux-Aylmer), QC J9H 3L5\n')
r_c_addr.font.size = Pt(9)
r_c_demo = p_c_rest.add_run('Démonstration en ligne : https://lecordondore.jopism.com  |  Date : Novembre 2026')
r_c_demo.font.size = Pt(9)
r_c_demo.font.bold = True

h1 = doc_comp.add_heading('1. Contexte & Diagnostic du Site Actuel', level=1)
h1.style.font.color.rgb = RGBColor(15, 23, 42)

doc_comp.add_paragraph('Le restaurant Le Cordon Doré propose une expérience culinaire distinctive et très recherchée dans le Vieux-Aylmer : la rencontre entre les techniques d\'un chef formé à Paris (25 ans d\'expérience), les saveurs authentiques de l\'Inde, un buffet hebdomadaire réputé et des viandes 100 % certifiées Halal.')
doc_comp.add_paragraph('Cependant, le site web actuel sous Wix (lecordondorerestaurantbistrot.ca) présente plusieurs freins majeurs qui réduisent vos ventes directes et compliquent l\'expérience des clients sur mobile :')

t_comp = doc_comp.add_table(rows=1, cols=3)
t_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
t_comp.autofit = False

h_cells_comp = t_comp.rows[0].cells
widths_comp = [Inches(2.1), Inches(2.2), Inches(2.7)]
headers_comp = ['Problème Actuel sur Wix', 'Impact sur vos Ventes', 'Solution dans le Nouveau Site']

for i, title in enumerate(headers_comp):
    h_cells_comp[i].text = title
    h_cells_comp[i].paragraphs[0].runs[0].font.bold = True
    h_cells_comp[i].paragraphs[0].runs[0].font.size = Pt(8.5)
    set_cell_background(h_cells_comp[i], 'F1F5F9')
    set_cell_margins(h_cells_comp[i], 80, 80, 100, 100)
    h_cells_comp[i].width = widths_comp[i]

comp_rows = [
    ('13 liens encombrés avec doublon et faute d\'orthographe ("Home" et "Acceuil" au lieu d\'"Accueil")', 'Les clients sur téléphone sont confus et quittent la page.', 'Menu épuré en 5 sections clés + bouton d\'appel direct visible en permanence.'),
    ('Pages fantômes indexées dans Google (copie-de-buffet-indien, copie-de-home)', 'Nuit à votre image de marque et à votre référencement local.', 'Structure SEO propre et optimisée sans aucune page dupliquée.'),
    ('Lien brisé & absence de combos', 'Le lien "Menu familiale Emporter" mène à "/copie-de-buffet-indien" qui n\'affiche qu\'une image de vacances sans texte.', 'Onglet dédié aux Combos Familiaux avec compositions claires et tarifs transparents.'),
    ('Menu Wix & fautes bilingues', 'Plat en doublon (Seekh kebab poulet affiché 2 fois) et fautes d\'anglais ("Scalops and calamary", "briyani rice").', 'Menu 100% vérifié, bilingue soigné, sans doublon ni faute d\'orthographe.'),
    ('Pages inachevées & rubriques vides', 'Page /blog affichant un post "Test" vide et /cave-cordon listant des catégories sans aucun prix.', 'Expérience fluide à 100% : cartes complètes des cocktails (14 $), vins et menus.'),
    ('3 numéros dispersés & faute dans le courriel ("cordondor1803" vs "cordondor18103")', 'Risque de courriels de réservation non reçus et confusion pour commander.', 'Dissociation claire : Appel direct commandes (819 682-5781), Texto réservations (819 661-5113) et courriel vérifié.'),
    ('Absence d\'un toggle bilingue fluide', 'Perte de clientèle anglophone à Aylmer et Ottawa.', 'Bascule FR / EN instantanée sur l\'ensemble du site en 1 clic.'),
    ('Mise en avant insuffisante du Buffet du mardi', 'Salle moins remplie pour l\'événement le plus rentable.', 'Bloc vedette dédié au Buffet Indien du mardi soir avec horaires et réservations.')
]

for row in comp_rows:
    r_cells = t_comp.add_row().cells
    for i, text in enumerate(row):
        r_cells[i].text = text
        p = r_cells[i].paragraphs[0]
        p.runs[0].font.size = Pt(8)
        set_cell_margins(r_cells[i], 60, 60, 90, 90)
        r_cells[i].width = widths_comp[i]
        if i == 1:
            p.runs[0].font.color.rgb = RGBColor(190, 18, 60)
        elif i == 2:
            p.runs[0].font.color.rgb = RGBColor(4, 120, 87)
            p.runs[0].font.bold = True

doc_comp.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 2: Audit éditorial
h2_audit = doc_comp.add_heading('2. Audit Éditorial : Coquilles & Incohérences Relevées sur le Site Wix', level=1)
h2_audit.style.font.color.rgb = RGBColor(15, 23, 42)

p_aud_intro = doc_comp.add_paragraph('Notre évaluation approfondie de votre site web actuel a mis en lumière plusieurs anomalies concrètes que la nouvelle version corrige intégralement :')
p_aud_intro.paragraph_format.space_after = Pt(2)

audit_bullets = [
    ('Plat en doublon au menu', 'Le plat « Seekh kebab poulet (2 pcs) » apparaît deux fois dans votre carte en ligne.'),
    ('Lien brisé et absence de combos', 'Le lien « Menu familiale Emporter » renvoie vers /copie-de-buffet-indien qui n\'affiche qu\'une vieille image d\'avis de vacances (CORDON VACANCES) au lieu des combos repas.'),
    ('Fautes d\'anglais sur les mets', 'Coquilles visibles par vos clients anglophones comme « Scalops and calamary » (au lieu de Scallops and calamari) et « briyani rice » (au lieu de biryani).'),
    ('Pages inachevées & rubriques vides', 'Votre page /blog affiche un article fictif « Test » sans aucun texte publié (« Aucun post publié dans cette langue »), et la page /cave-cordon liste des catégories d\'alcool sans aucun met ni prix.'),
    ('Faute dans la navigation principale', 'Doublon entre « Home » et « Acceuil » (avec faute d\'orthographe sur le mot Accueil).'),
    ('Adresses courriel contradictoires', 'cordondor1803@gmail.com affiché en bas de page contre cordondor18103@gmail.com sous Réservation (risque d\'échec de réception).'),
    ('Pages fantômes indexées par Google', 'Pages de test créées par Wix non nettoyées (copie-de-buffet-indien, menus-1, copie-de-home) qui nuisent à votre crédibilité et référencement.')
]

for title, desc in audit_bullets:
    p = doc_comp.add_paragraph()
    p.paragraph_format.space_after = Pt(1.5)
    r_pt = p.add_run(f'• {title} : ')
    r_pt.bold = True
    r_pt.font.color.rgb = RGBColor(190, 18, 60)
    p.add_run(desc)

doc_comp.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 3: Les 5 Piliers
h3_pil = doc_comp.add_heading('3. Les 5 Piliers Stratégiques de la Nouvelle Version', level=1)
h3_pil.style.font.color.rgb = RGBColor(15, 23, 42)

points_comp = [
    ('[Pilier 1] Vitesse & Mobile-First (< 1 seconde)', 'Plus de 85 % des recherches de restaurants se font sur téléphone. La nouvelle page se charge instantanément sans lenteurs de constructeur.'),
    ('[Pilier 2] Commandes Téléphoniques Directes', 'Un bouton d\'appel en 1-clic incite vos clients à commander directement au restaurant plutôt que via les applications de livraison tierces (comme DoorDash), vous épargnant jusqu\'à 30 % de commissions.'),
    ('[Pilier 3] Valorisation des Signes de Confiance', 'Affichage immédiat en tête de page de vos viandes 100 % certifiées Halal, de la note Google de 4.8★ et de la formation prestigieuse à Paris de votre Chef formé à Paris.'),
    ('[Pilier 4] Combos Familiaux & Buffet Vedette', 'Mise en lumière claire du Buffet Indien du mardi soir et des festins à emporter (2 à 6 personnes) pour augmenter le panier moyen.'),
    ('[Pilier 5] Service Traiteur & Cartes-Cadeaux', 'Un espace dédié avec demande de soumission directe pour capter les mariages, fêtes privées et dîners corporatifs dans l\'Outaouais.')
]

for title, desc in points_comp:
    p = doc_comp.add_paragraph()
    r_pt = p.add_run(f'• {title} : ')
    r_pt.bold = True
    r_pt.font.color.rgb = RGBColor(180, 83, 9)
    p.add_run(desc)

doc_comp.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 4: PORTRAIT FINANCIER
h4_fin = doc_comp.add_heading('4. Portrait financier : frais actuels vs nouvelle approche', level=1)
h4_fin.style.font.color.rgb = RGBColor(15, 23, 42)

fin_points = [
    ('Commissions des applications de livraison tierces (comme DoorDash)', 'Au Canada, les plateformes prélèvent généralement de 20 % à 30 % sur les livraisons et de 8 % à 15 % sur les commandes à emporter. Sur une commande moyenne de 40 $, cela représente 8 $ à 12 $ de marge perdue pour votre restaurant.'),
    ('Frais actuels d\'hébergement Wix au Canada', 'L\'abonnement Wix représente un coût récurrent d\'environ 250 $ à 530 $+ CAD par an pour héberger le site actuel.'),
    ('Risque commercial sur les réservations', 'Deux adresses Gmail divergentes sont actuellement affichées sur votre page d\'accueil (cordondor1803@gmail.com vs cordondor18103@gmail.com), risquant de vous faire perdre des réservations par courriel.'),
    ('Nouvelle approche : Hébergement moderne à 0 $ / mois', 'En migrant vers une architecture web statique ultra-performante, vos frais mensuels d\'hébergement passent à 0 $. Vous ne conservez que le coût de base annuel de votre nom de domaine .ca (~15 $ à 25 $/an).'),
    ('Mon offre clé en main gratuite', 'Je réalise gratuitement la conception du nouveau site, la migration technique et le transfert de lecordondorerestaurantbistrot.ca avec zéro interruption de service pour vos clients, tout en assurant la maintenance continue (mise à jour des prix, menus de saison, horaires du buffet/fêtes).'),
    ('Option Phase 2 : Commande en ligne DoorDash Storefront (0 $/mois, 0 % de commission)', '0 $ de frais mensuels et 0 % de commission marketplace sur les commandes passées depuis votre site web (seuls les frais bancaires de carte s\'appliquent). Zéro friction pour Le Cordon Doré : vous êtes déjà sur DoorDash, donc les commandes se synchronisent automatiquement avec votre menu existant et sonnent directement sur la même tablette que vous avez déjà au comptoir ! Les Dashers assurent la livraison à domicile sans la commission de 20 % à 29 %.'),
    ('Redirection de courriel professionnel gratuit (optionnel)', 'Possibilité d\'ajouter sans frais une redirection de courriel professionnel (info@lecordondorerestaurantbistrot.ca) pour professionnaliser vos échanges traiteur et réservations.')
]

for title, desc in fin_points:
    p = doc_comp.add_paragraph()
    r_pt = p.add_run(f'• {title} : ')
    r_pt.bold = True
    p.add_run(desc)

doc_comp.add_paragraph().paragraph_format.space_after = Pt(4)

# Section 5: PLAN DE DÉPLOIEMENT
h5_plan = doc_comp.add_heading('5. Plan de Déploiement Simple & Sans Interruption', level=1)
h5_plan.style.font.color.rgb = RGBColor(15, 23, 42)

steps_comp = [
    ('Étape 1 : Validation du Prototype en ligne', 'Vous testez la démonstration complète sur votre mobile à l\'adresse https://lecordondore.jopism.com.'),
    ('Étape 2 : Ajustement des Menus & Tarifs', 'Mise à jour immédiate des prix, plats saisonniers ou coordonnées selon vos consignes exactes.'),
    ('Étape 3 : Bascule Transparente vers Votre Domaine', 'Connexion directe à votre adresse lecordondorerestaurantbistrot.ca en moins de 2 heures, sans aucune interruption pour vos clients.'),
    ('Bénéfice Immédiat : Économies & Sérénité', 'Zéro mensualité Wix, site ultra-rapide et maintenance technique prise en charge.')
]

for title, desc in steps_comp:
    p = doc_comp.add_paragraph()
    r_st = p.add_run(f'✓ {title} — ')
    r_st.bold = True
    p.add_run(desc)

add_contact_block(doc_comp)
doc_comp.save(r'C:\Users\joppa\Desktop\Documents\GitHub\Le Cordon Doré\Proposition-Strategique-Complete.docx')
print('Proposition-Strategique-Complete.docx updated successfully')
