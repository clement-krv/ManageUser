<p align="center"><a href="https://www.laboulangere.com/" target="_blank">
    <img src="https://www.laboulangere.com/app/uploads/2022/07/logo-la-boulangere.png" alt="La Boulangère Logo">
</a></p>

# Manage User

Ce projet a été développé pour La Boulangère. Il contient des scripts pour gérer les utilisateurs Active Directory et Azure Active Directory. 

---
## Description

J'ai dû le développer pour optimiser les données, et également pour éviter les anomalies lors des saisies spécifiques. Le code est anonyme, c'est-à-dire que toutes les données sensibles ont été remplacées. Le code est voué à être amélioré avec une interrogation sur une API pour récupérer automatiquement les données saisies par le service RH.

---
## Scripts

- [`createAdAzure.ps1`](assets/scripts/createAdAzure.ps1) : Ce script permet de créer un utilisateur Active Directory avec ou sans synchronisation Azure.
- [`createAzure.ps1`](assets/scripts/createAzure.ps1) : Ce script permet de créer un compte Azure.
- [`deleteAdAzure.ps1`](assets/scripts/deleteAdAzure.ps1) : Ce script permet de supprimer un utilisateur Active Directory avec ou sans conservation de la boîte mail en partagée.
- [`menu.ps1`](assets/scripts/menu.ps1) : Ce script affiche un menu pour choisir l'action à effectuer.
- [`ManageUser.ps1`](ManageUser.ps1) : Ce script principal fait appel aux autres scripts.

---
## Utilisation

Pour utiliser ces scripts, ouvrez PowerShell et naviguez jusqu'au répertoire contenant les scripts. Ensuite, exécutez le script `ManageUser.ps1`.

---
## Crédits

Ce projet a été créé par [Clément Kerviche](https://github.com/clement-krv) pour [La Boulangère](https://www.laboulangere.com).