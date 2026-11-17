📊 Call Center Operations & SLA Recovery Dashboard (Excel)

**Auteur :** Kodjo Godwin BADA  
**Outils utilisés :** Microsoft Excel (Tableaux Croisés Dynamiques, Pivot Charts, Slicers, Timeline, `GETPIVOTDATA`)  
**Dataset :** Call Centre Queue Simulation (51 708 appels simulés sur 1 an)[cite: 5]

![Aperçu du Dashboard](./Capture%20d'écran%202026-11-17%20133417.png)

---

## 🎯 1. Contexte & Objectifs Métier

Ce projet analyse l'activité annuelle d'un centre d'appels opérant du **lundi au vendredi, de 8h00 à 18h00**, avec une équipe fixe de **4 agents** en ligne et un temps moyen de traitement (DMT) théorique de **5 minutes (300 secondes)**.

* **Objectif opérationnel (SLA) :** Répondre à au moins **90 % des appels en moins de 60 secondes**.
* **Problématique :** Identifier les créneaux horaires, les jours de la semaine et les mois où la performance chute sous la barre critique des 90 % afin de proposer des ajustements d'effectifs.

---

## 🏗️ 2. Architecture & Processus de Construction

Le rapport repose sur une architecture **100 % dynamique pilotée par Tableaux Croisés Dynamiques (TCD)** sans recours au VBA, organisée en trois couches distinctes :

### Étape 1 : Préparation et Nettoyage des Données (`Data`)
* **Standardisation temporelle :** Extraction de l'heure d'arrivée (`Hours`), du jour de la semaine (`Jour`), du mois (`Months`) et du numéro de semaine (`Numéro de Semaine`) à partir des horodatages bruts[cite: 5].
* **Conversion des durées (Résolution du piège temporel Excel) :** Excel stockant le temps en jours ($1 = 24\text{h} = 86\,400\text{s}$), la colonne `service_length` (en secondes) a été divisée par `86400` pour permettre un affichage précis au format personnalisé `m:ss` (ex. `4:54`) sans fausser les moyennes.
* **Indicateur SLA :** Création d'un indicateur logique (`meets_standard` : `TRUE` / `FALSE`) évaluant si le temps d'attente est $\le 60\text{ secondes}$[cite: 5].

### Étape 2 : Modélisation Analytique (`Pivot_Data`)
Construction de 4 Tableaux Croisés Dynamiques dédiés pour alimenter chaque bloc visuel :
1. **TCD 1 (Performance Horaire) :** Volume d'appels (`Count`) et taux de SLA (`Average`) par heure (`08h` à `17h`)[cite: 5].
2. **TCD 2 (Charge Hebdomadaire) :** Délai Moyen d'Attente (`DMA` en secondes) et volume d'appels par jour ouvré (`Lundi` à `Vendredi`)[cite: 5].
3. **TCD 3 (Tendance Annuelle) :** Évolution mensuelle du Service Level (`SL`) comparée à la cible fixe de `90%` (`janv` à `déc`)[cite: 5].
4. **TCD 4 (Capacité & Traitement) :** Durée Moyenne de Traitement (`DMT`) par heure comparée au standard cible de `5:00` (`Target of DMT`)[cite: 5].

### Étape 3 : Cartes KPI Dynamiques (`GETPIVOTDATA`)
Les formes arrondies en en-tête sont reliées dynamiquement aux TCD via des cellules intermédiaires utilisant la fonction `GETPIVOTDATA` (Excel n'acceptant pas les formules imbriquées directement dans les objets graphiques) :
* **Target Actuel (SLA) :** `91,83%`[cite: 5]
* **DMA (Délai Moyen d'Attente) :** `18s`[cite: 5]
* **Appel Total :** `51 708`[cite: 5]
* **Écart vs Target (90%) :** `+1,83%`[cite: 5]

### Étape 4 : Design UI/UX & Interactivité (`Dashboard`)
* **Thème Dark Mode :** Fond sombre (`#1E2022` / `#282A2E`) avec bordures arrondies et accents cyan (`#00D2FF`) et orange (`#FF6D00`) pour un contraste optimal[cite: 5].
* **Contrôles Universels :** Création de styles de segments (*Slicers*) personnalisés pour le **Numéro de Semaine** (grille multi-colonnes), le **Jour**, le statut **`meets_standard`**, ainsi qu'une **Timeline** interactive sur le champ `date`, tous reliés à l'ensemble des TCD via *Connexions de rapport*[cite: 5].

---

## 🔍 3. Principaux Enseignements & Recommandations

L'analyse visuelle met en évidence trois phénomènes opérationnels majeurs :

1. **Dégradation progressive au cours de l'année :** Bien que le SLA annuel global atteigne **91,83 %**, la courbe mensuelle montre une érosion constante passant de **> 95 % au T1** à une chute sous le seuil critique des **90 % dès le mois d'octobre** (atteignant ~83 % en décembre)[cite: 5].
2. **Goulot d'étranglement à la pause méridienne (13h00) :** Le graphique combiné horaire révèle que le volume reste élevé à **13h** (> 5 200 appels cumulés) tandis que le SLA chute brutalement à **90,0 %** (contre > 92 % le matin), traduisant une sous-capacité liée aux pauses déjeuners des agents[cite: 5].
3. **Pic de charge le Vendredi :** Le **vendredi** concentre le volume d'appels le plus élevé (**10 695 appels**) et le temps d'attente moyen le plus long (**21 secondes** contre 14s le mercredi)[cite: 5].

### 💡 Plan d'action recommandé :
* **Décaler les pauses déjeuners** entre 12h00 et 14h30 afin de maintenir 4 agents actifs durant le pic de 13h00[cite: 5].
* **Renforcer l'équipe le vendredi** (+1 agent en renfort ou réallocation de tâches back-office) pour absorber le pic de 10 695 appels[cite: 5].
* **Anticiper la saisonnalité du T4** (octobre à décembre) par un recrutement temporaire ou des heures supplémentaires ciblées dès septembre[cite: 5].

---

## 🚀 4. Comment utiliser ce fichier

1. Téléchargez le fichier `.xlsx` depuis ce dépôt.
2. Ouvrez-le dans **Microsoft Excel** (version 2016 ou supérieure recommandée pour le support complet des Timelines et styles personnalisés).
3. Utilisez les **Slicers** à droite (Semaine, Jour, Statut SLA) ou faites glisser la **Timeline** pour filtrer l'ensemble des graphiques et des cartes KPI en temps réel[cite: 5].
