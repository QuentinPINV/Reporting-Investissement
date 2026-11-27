/*
 * Paramètres du site « Reporting Service Investissement ».
 * Seules les deux premières valeurs sont à remplir, avec ce que donne le service informatique
 * après la création de l'inscription d'application dans Microsoft Entra.
 */
window.EURYALE_CONFIG = {
  // ID d'application (client) de l'inscription Entra — à remplir
  clientId: 'A_REMPLIR',

  // ID de l'annuaire (locataire) Entra — à remplir, ou laisser le domaine de l'entreprise
  tenant: 'euryale.com',

  // Bibliothèque « Documents partagés » du site SharePoint Equipe Invest
  driveId: 'b!UtXL-qZoMkacrvfg72ujHU9aCIeOe4dArjkjJWeHNdGFqTShNbJ2QrvrbgaGSFRQ',

  // Fichiers lus à chaque ouverture (chemin depuis la racine de la bibliothèque) et tableaux Excel utilisés
  dealflow: {
    path: 'Reporting Service Investissement/Fichier reporting_Dealflow_2026.xlsx',
    table: 'Tableau1',          // onglet « Dealflow - Synthèse »
    resumeSheet: 'Résumé',      // onglet contenant le tableau « Devise / Taux → EUR »
  },
  mouvements: {
    path: 'Reporting Service Investissement/Fichier reporting_Mouvements Patrimoine_2026.xlsx',
    engagements: 'Tableau3',    // onglet « 1) ENGAGEMENTS »
    acquisitions: 'Tableau2',   // onglet « 2) ACQUISITIONS »
    cessions: 'Tableau15',      // onglet « 3) CESSIONS »
  },
};
