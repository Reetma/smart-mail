// Message type : objet et texte déjà en place dans le panneau, modifiables.
// Repris quand le champ est vide à l'ouverture, et par le bouton « Message type ».
// Pour changer le message type (un lien, une date…), modifier ce fichier.
//
// Exposé dans window.PP.Modele.

(function (root) {
  "use strict";
  var PP = (root.PP = root.PP || {});

  var LIEN = "https://reetma.com/ecole";

  PP.Modele = {
    objet: "Certification IA - Avant 31 Oct.",
    html: [
      '<div style="font-family: Aptos, sans-serif; font-size: 16px; color: rgb(0, 0, 0);">',
      "<div>Chers étudiants,</div>",
      "<div><br></div>",
      "<div>Les compétences en intelligence artificielle figurent parmi les plus recherchées par les employeurs : vous devez vous inscrire avant le <strong>31 octobre</strong>.</div>",
      "<div><br></div>",
      "<div>Vous pourrez ajouter votre certificat de réussite Reetma directement à votre profil LinkedIn et à votre CV, offrant ainsi aux recruteurs une preuve concrète de votre expérience pratique en IA.</div>",
      "<div><br></div>",
      '<div><strong><a href="' + LIEN + '">OBTENEZ VOTRE CERTIFICATION AETHER EN IA</a></strong></div>',
      "<div><br></div>",
      "<div>71&nbsp;% des dirigeants préféreraient recruter une personne moins expérimentée possédant des compétences en IA plutôt qu’une personne plus expérimentée qui en serait dépourvue. Les places sont limitées.</div>",
      "<div><br></div>",
      '<div>Réservez votre place et commencez la certification → <a href="' + LIEN + '">' + LIEN + "</a></div>",
      "<div><br></div>",
      "<div>Alors que l’IA prend une importance croissante dans tous les secteurs, mettre en avant ces compétences peut considérablement renforcer votre candidature à un stage ou à un premier emploi après vos études.</div>",
      "<div><br></div>",
      "<div>Cordialement,<br>Reetma School</div>",
      "</div>",
    ].join(""),
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
