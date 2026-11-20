Plaquette commerciale Eulidia : sept slides figées à insérer telles quelles juste avant la slide de fin, systématiquement dans une proposition commerciale et, ailleurs, seulement quand l'utilisateur la demande explicitement.

**Groupe** : Plaquette commerciale · **Disposition** : 05 Slide vide claire (fonds dégradés pleine page posés sur la slide) · **Source** : `eulidia-socle-2026.pptx` (section *Socle et plaquette*)

| Ordre | Slide | Contenu | Vignette |
| --- | --- | --- | --- |
| 1 | ChiffresCles | Positionnement d’Eulidia : titre, trois piliers, accroche et quatre chiffres clés. | `slide-39.jpg` |
| 2 | MurLogosClients | Références clients : accroche, titre « Just Happy Clients » et mur de logos clients. | `slide-40.jpg` |
| 3 | OffreEndToEnd | Offre de bout en bout : Stratégie, Delivery, Adoption. | `slide-41.jpg` |
| 4 | PartenairesTechno | Partenariats éditeurs : constellation de logos. | `slide-42.jpg` |
| 5 | FocusPartenaire | Focus sur un éditeur partenaire (statut, chiffres, certifications). | `slide-43.jpg` |
| 6 | EquipeDirection | Équipe de direction : trois directeurs et leurs domaines. | `slide-44.jpg` |
| 7 | EulidiaInstitute | Offre de formation Eulidia Institute. | `slide-45.jpg` |

- **Quand** : toujours dans une proposition commerciale ; dans tout autre document (générique, kick-off, comité projet…), uniquement si l'utilisateur la demande explicitement. Sans demande, la retirer du socle avec `retirer_plaquette()`.
- **Où** : toujours les 7, dans cet ordre, **après le contenu et avant la SlideFin** (`placer_plaquette()`) : … dernière slide de contenu → plaquette (7 slides) → SlideFin.
- Elles sont déjà présentes dans le socle (skill `eulidia-powerpoint`) : ne pas les reconstruire, ne pas les copier depuis une image, ne modifier ni texte, ni logo, ni photo.
- Ne pas : réutiliser leurs mises en page (fonds dégradés pleine page, grands titres 54 pt, cartes translucides) pour construire d'autres slides ; la disposition 05 *Slide vide claire* n'est utilisée que par elles.
- Ne pas : les numéroter dans un sommaire ni leur donner une section propre.
