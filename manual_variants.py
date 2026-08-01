"""
manual_variants.py
==================
Handmatig gedefinieerde maat-varianten voor Brokken-producten.

Voor elk Brokken-product hebben we 2 korreltypes (zoals op yoggies.cz):
- Normaal: grotere korrels voor middelgrote en grote rassen
- Mini:    kleinere korrels voor kleine rassen en puppy's

Op de productpagina verschijnt een dropdown "Kies grootte". De "Bestel"-knop
linkt naar de juiste DogLife-pagina per selectie. Als een maat NIET op DogLife
staat (lege URL), verschijnt automatisch het contactblok (telefoon + e-mail).

URL's gevalideerd op doglife.be (juni 2026).
"""

MANUAL_VARIANTS = {

    "Yoggies brokken Active Eend & Wild": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/active-eend-en-hert/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/active-eend-en-hert-mini/"},
    ],

    "Yoggies brokken Lam & Witvis": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/lam-en-witte-vis/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/lam-en-witte-vis-mini/"},
    ],

    "Yoggies brokken Geit monoprote\u00b4ne": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/hypoalergeen-geitenvlees-met-groenten/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/hypoalergeen-geitenvlees-met-groenten-mini/"},
    ],

    "Yoggies brokken Kalkoen monoprote\u00b4ne": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/hypoalergeen-kalkoenvlees-met-gierst-en-hennepolie-klassiek/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/hypoalergeen-kalkoenvlees-met-gierst-en-hennepolie-mini/"},
    ],

    "Yoggies brokken Kip & Rund": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/kip-en-rund-klassiek/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/kip-en-rund-mini/"},
    ],

    "Yoggies brokken Varkensvlees Ib\u00dArico monoprote\u00b4ne": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/iberico-met-appels-hypoallergeen/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/iberico-met-appels-mini-hypoallergeen/"},
    ],

    "Yoggies brokken VET Gastro Sensitive": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/vet-gastro-sensitive-met-kalkoen-en-probiotica-5-kg/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/vet-gastro-sensitive-mini-met-kalkoen-en-probiotica-5-kg/"},
    ],

    "Yoggies brokken VET Insect Derma": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/vet-insect-derma-met-insecteneiwit-met-probiotica-5-kg/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/vet-insect-derma-mini-met-insecteneiwit-en-probiotica-5-kg/"},
    ],

    "Yoggies brokken VET Veggie vleesvrij": [
        {"size": "Normaal", "dl_url": "https://doglife.be/product/vet-veggie-zonder-vlees-met-probiotica-5kg/"},
        {"size": "Mini",    "dl_url": "https://doglife.be/product/vet-veggie-mini-zonder-vlees-met-probiotica-5kg/"},
    ],
}