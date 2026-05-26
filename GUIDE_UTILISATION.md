# 🍽️ Application Excel de Gestion de Restaurant

## 📋 Présentation

Application Excel professionnelle **clé en main** pour la gestion complète d'un restaurant, développée **sans VBA**. L'application utilise les fonctionnalités avancées d'Excel : formules avancées, tableaux structurés, mise en forme conditionnelle et tableaux de bord dynamiques.

---

## 📁 Structure du Fichier

**Fichier :** `GestionRestaurant.xlsx`

### 13 Feuilles Principales :

| # | Feuille | Description |
|---|---------|-------------|
| 1 | **ACCUEIL** | Page d'accueil avec navigation et alertes |
| 2 | **DASHBOARD** | Tableau de bord principal avec KPI |
| 3 | **PRODUITS** | Gestion des produits et matières premières |
| 4 | **CATEGORIES** | Catégories et sous-catégories de produits |
| 5 | **STOCKS** | Suivi des mouvements de stock (entrées/sorties) |
| 6 | **RECETTES** | Gestion des recettes et calcul des coûts matières |
| 7 | **INGREDIENTS** | Détail des ingrédients par recette |
| 8 | **VENTES** | Enregistrement des ventes et chiffre d'affaires |
| 9 | **FOURNISSEURS** | Gestion des fournisseurs et commandes |
| 10 | **INVENTAIRES** | Contrôles périodiques et détection d'écarts |
| 11 | **TABLEAUX_BORD** | Tableaux de bord analytiques détaillés |
| 12 | **PARAMETRES** | Configuration et paramètres système |
| 13 | **AIDE** | Guide d'utilisation et bonnes pratiques |

---

## ⚡ Fonctionnalités Clés

### 🔔 Alertes Automatiques
- **⚠️ RUPTURE** : Stock = 0
- **⚠️ CRITIQUE** : Stock < Stock Minimum
- **⚡ ATTENTION** : Stock < Stock de Sécurité
- **✓ OK** : Stock normal

### 📊 Calculs Automatiques
- Coût matière par recette
- Marge brute et taux de marge
- Valeur totale du stock
- Panier moyen
- Rotation des stocks
- Détection des écarts d'inventaire

### 📈 Tableaux de Bord
- Ventes totales et marges
- Nombre de transactions
- Produits les plus vendus
- Alertes en temps réel
- Performance par catégorie

---

## 🚀 Prise en Main Rapide

### Ordre de Configuration Recommandé

1. **CATEGORIES** → Créez vos catégories et sous-catégories
2. **PRODUITS** → Ajoutez vos produits avec prix et seuils
3. **FOURNISSEURS** → Enregistrez vos fournisseurs
4. **RECETTES** → Créez vos plats et menus
5. **INGREDIENTS** → Associez les ingrédients aux recettes
6. **STOCKS** → Enregistrez les mouvements initiaux
7. **VENTES** → Saisissez les ventes quotidiennes
8. **DASHBOARD** → Consultez les indicateurs

---

## 📝 Formules Principales Utilisées

### Statut du Stock (feuille PRODUITS)
```excel
=IF(G2=0,"⚠️ RUPTURE",IF(G2<I2,"⚠️ CRITIQUE",IF(G2<J2,"⚡ ATTENTION","✓ OK")))
```

### Coût Matière d'une Recette (feuille RECETTES)
```excel
=SUMIF(INGREDIENTS!A:A,A3,INGREDIENTS!H:H)
```

### Marge Brute (feuille RECETTES)
```excel
=Prix_Vente - Coût_Matière
```

### Taux de Marge (feuille RECETTES)
```excel
=IFERROR(Marge/Prix_Vente*100,0)
```

### Total Vente (feuille VENTES)
```excel
=Prix_Unitaire * Quantité
```

### Écart d'Inventaire (feuille INVENTAIRES)
```excel
=Stock_Réel - Stock_Théorique
```

---

## 💡 Bonnes Pratiques

### Quotidien
- ✅ Vérifier les alertes de stock sur la page ACCUEIL
- ✅ Enregistrer toutes les ventes dans VENTES
- ✅ Noter les mouvements de stock (pertes, ajustements)

### Hebdomadaire
- ✅ Faire un inventaire partiel des produits critiques
- ✅ Analyser les marges par plat
- ✅ Vérifier les commandes fournisseurs en attente

### Mensuel
- ✅ Réaliser un inventaire complet
- ✅ Analyser les performances globales
- ✅ Ajuster les seuils de stock si nécessaire
- ✅ Sauvegarder le fichier

---

## 🔧 Maintenance

### Actualiser les Données
Dans Excel : **Données** → **Actualiser tout**

### Protéger les Feuilles
1. **Révision** → **Protéger la feuille**
2. Définir un mot de passe
3. Sélectionner les actions autorisées

### Sauvegarde
- **Fichier** → **Enregistrer sous** → Ajouter la date
- Exemple : `GestionRestaurant_2024-01-15.xlsx`

---

## 🎨 Design et Ergonomie

- **Couleurs professionnelles** : Bleu foncé, bleu moyen, accents colorés
- **Codes visuels** : Icônes pour les statuts (⚠️, ⚡, ✓)
- **Tableaux structurés** : Mise en forme automatique avec lignes alternées
- **Navigation intuitive** : Page d'accueil avec liens vers tous les modules

---

## 📊 Indicateurs de Performance (KPI)

### Financiers
- Chiffre d'affaires total
- Marge brute totale
- Taux de marge moyen
- Panier moyen

### Stocks
- Valeur totale du stock
- Nombre de produits critiques
- Rotation des stocks
- Couverture en jours

### Opérationnels
- Nombre de ventes
- Recettes actives
- Nombre de fournisseurs
- Dernier inventaire

---

## ⚠️ Limitations et Recommandations

### Sans VBA
Cette application n'utilise **pas de VBA**, ce qui garantit :
- ✅ Compatibilité maximale
- ✅ Sécurité accrue (pas de macros)
- ✅ Facilité de déploiement

### Recommandations
- Utiliser Excel 2016 ou version ultérieure
- Activer les mises à jour automatiques des formules
- Ne pas supprimer les colonnes de formules
- Tester les formules après toute modification structurelle

---

## 🆘 Dépannage

### Les formules ne se mettent pas à jour
- Vérifier que le calcul automatique est activé : **Formules** → **Options de calcul** → **Automatique**

### Les tableaux structurés ne s'étendent pas
- Ajouter les nouvelles données directement sous la dernière ligne du tableau
- Le tableau devrait s'étendre automatiquement

### Alertes non affichées
- Vérifier que les colonnes de statut contiennent bien les formules
- Contrôler les références de cellules dans les formules

---

## 📞 Support et Améliorations

### Pour ajouter de nouvelles fonctionnalités
1. Dupliquer une feuille existante comme modèle
2. Adapter les formules aux nouveaux besoins
3. Tester soigneusement avant utilisation en production

### Personnalisation
- Modifier les couleurs dans l'onglet **Accueil** → **Mise en forme**
- Ajuster les largeurs de colonnes selon vos besoins
- Ajouter des graphiques via **Insertion** → **Graphiques recommandés**

---

## 📄 Licence et Utilisation

Cette application est fournie **gratuitement** pour un usage professionnel ou personnel.

### Droits
- ✅ Utilisation commerciale autorisée
- ✅ Modification autorisée
- ✅ Distribution autorisée

### Restrictions
- ❌ Revente en l'état interdite
- ❌ Revendication de paternité interdite

---

## 🎯 Conclusion

Cette application Excel offre une solution complète et professionnelle pour gérer tous les aspects de votre restaurant :

- 📦 **Stocks** optimisés avec alertes automatiques
- 💰 **Ventes** suivies avec calcul de marges
- 📖 **Recettes** maîtrisées avec coûts précis
- 🚚 **Fournisseurs** organisés
- 📊 **Tableaux de bord** pour piloter votre activité

**Sans VBA, 100% Excel, prêt à l'emploi !**

---

© 2024 - Application de Gestion de Restaurant v1.0
