'use strict';

/* ── DOM helpers ── */
const g   = id => document.getElementById(id);
const val = id => parseFloat(g(id).value) || 0;
const radio   = name => { const el = document.querySelector(`input[name="${name}"]:checked`); return el ? el.value : null; };
const checked = id   => g(id).checked;
const fmt     = n    => Math.round(n) + ' g';

function showResult(prefix) {
  g(prefix + '-placeholder').style.display = 'none';
  g(prefix + '-result').classList.add('visible');
}

/* ════════════════════════════════════════════════════════════
   CORE FORMULAS
   Source: NRC 2006 "Nutrient Requirements of Dogs and Cats"
           FEDIAF "Nutritional Guidelines" (2021)
════════════════════════════════════════════════════════════ */

/**
 * RER — Resting Energy Requirement
 * Formula: 70 × W(kg)^0.75   [NRC 2006, both species]
 * Returns kcal/day
 */
function rer(kg) {
  return 70 * Math.pow(kg, 0.75);
}

/**
 * Dog MER factor
 * MER = RER × factor
 * Factors from FEDIAF 2021 Table 2 & NRC 2006 Ch.3
 *
 * Life stage × Activity × BCS × Reproductive
 */
function dogMerFactor(lifeStage, activity, bcs, castrated, pregnant) {
  // Base MER factors (life stage + activity)
  const tbl = {
    //             low    normal  high   xhigh
    puppy:  { low:2.5,  normal:3.0, high:3.5, xhigh:4.0 },
    adult:  { low:1.4,  normal:1.6, high:2.0, xhigh:2.5 },
    senior: { low:1.1,  normal:1.2, high:1.4, xhigh:1.6 },
  };
  const stage  = (lifeStage === 'junior') ? 'adult' : (lifeStage || 'adult');
  const actKey = activity || 'normal';
  let factor   = (tbl[stage] || tbl.adult)[actKey] || 1.6;

  // BCS correction (body condition score)
  if (bcs === 'thin') factor *= 1.10;  // underweight: slightly more
  if (bcs === 'over') factor *= 0.80;  // overweight: restrict

  // Reproductive corrections (highest priority)
  if (pregnant)          factor = Math.max(factor, 2.0); // gestating/lactating ~2.0–3.0×RER
  else if (castrated)    factor *= 0.85; // −15% for neutered

  return Math.round(factor * 100) / 100;
}

/**
 * Cat MER factor
 * FEDIAF 2021 Table 5 for cats
 */
function catMerFactor(lifeStage, lifestyle, bcs, castrated, pregnant) {
  const tbl = {
    //            indoor  outdoor
    kitten: { indoor:2.5, outdoor:3.0 },
    adult:  { indoor:1.0, outdoor:1.4 },  // indoor neutered ~1.0; intact outdoor ~1.4
    senior: { indoor:0.8, outdoor:1.0 },
  };
  const stage  = lifeStage  || 'adult';
  const lstyle = lifestyle   || 'outdoor';
  let factor   = (tbl[stage] || tbl.adult)[lstyle] || 1.2;

  if (bcs === 'thin') factor *= 1.10;
  if (bcs === 'over') factor *= 0.80;

  if (pregnant)        factor = Math.max(factor, 2.0);
  else if (castrated)  factor *= 0.80; // −20% for neutered cats (higher insulin sensitivity)

  return Math.round(factor * 100) / 100;
}

/**
 * Dog BARF % of body weight
 * Standard BARF guidelines (Ian Billinghurst / Dogs Naturally / BARF World)
 * Adult: 2–3% BW/day ; Puppy: 4–6% BW/day ; Senior: 1.5–2%
 */
function barfPctDog(lifeStage, activity, bcs, castrated) {
  const tbl = {
    //              low   normal  high
    puppy:  { low:4.0,  normal:5.0, high:6.0 },
    adult:  { low:2.0,  normal:2.5, high:3.0 },
    senior: { low:1.5,  normal:2.0, high:2.5 },
  };
  const stage  = lifeStage || 'adult';
  const actKey = activity  || 'normal';
  let pct = (tbl[stage] || tbl.adult)[actKey] || 2.5;

  if (bcs === 'thin') pct *= 1.10;
  if (bcs === 'over') pct *= 0.85;
  if (castrated)      pct *= 0.90;

  return Math.round(pct * 100) / 100;
}

/**
 * Cat BARF % of body weight
 */
function barfPctCat(lifeStage, lifestyle, bcs, castrated) {
  const tbl = {
    kitten: { indoor:4.0, outdoor:5.0 },
    adult:  { indoor:2.0, outdoor:2.5 },
    senior: { indoor:1.5, outdoor:2.0 },
  };
  const stage  = lifeStage || 'adult';
  const lstyle = lifestyle  || 'outdoor';
  let pct = (tbl[stage] || tbl.adult)[lstyle] || 2.5;

  if (bcs === 'thin') pct *= 1.10;
  if (bcs === 'over') pct *= 0.85;
  if (castrated)      pct *= 0.80;

  return Math.round(pct * 100) / 100;
}

/** Recommended meals/day */
function mealsPerDay(lifeStage) {
  if (lifeStage === 'puppy' || lifeStage === 'kitten') return 3;
  return 2;
}

/* ════════════════════════════════════════════════════════════
   CALCULATOR FUNCTIONS
════════════════════════════════════════════════════════════ */

/* 1. DOG – DROOGVOER (BROKKEN) */
function calcDogBrok() {
  const kg          = val('db-weight');
  const kcalPer100  = val('db-kcal');
  const life        = radio('db-life');
  const act         = radio('db-act');
  const bcs         = radio('db-bcs');
  const cast        = checked('db-castrated');
  const preg        = checked('db-pregnant');

  if (!kg || kg < 0.5)           { alert('Vul een geldig gewicht in (minimaal 0,5 kg).'); return; }
  if (!kcalPer100 || kcalPer100 < 100) { alert('Vul een geldige caloriewaarde in.'); return; }

  const RER    = rer(kg);
  const factor = dogMerFactor(life, act, bcs, cast, preg);
  const MER    = RER * factor;
  const grams  = (MER / kcalPer100) * 100;
  const meals  = mealsPerDay(life);

  g('db-grams').textContent    = Math.round(grams);
  g('db-kcal-out').textContent = `≈ ${Math.round(MER)} kcal / dag`;
  g('db-meals').textContent    = meals + ' maaltijden';
  g('db-per-meal').textContent = Math.round(grams / meals) + ' g per maaltijd';
  g('db-factor').textContent   = factor + '× RER';
  g('db-rer').textContent      = Math.round(RER) + ' kcal / dag';

  const notes = [];
  if (preg)          notes.push('Drachtige/zogende honden: verhoog geleidelijk en volg gewicht wekelijks op.');
  if (bcs === 'over') notes.push('Overgewicht: controleer gewicht om de 2 weken en pas portie aan als er geen daling is.');
  if (bcs === 'thin') notes.push('Ondergewicht: verhoog portie geleidelijk; pas terug zodra ideaalgewicht bereikt is.');
  if (life === 'puppy') notes.push('Pups groeien snel — weeg wekelijks en pas de portie steeds aan naar het actuele lichaamsgewicht.');
  g('db-note').textContent = notes.join(' ') || 'Dit is een richtlijn. Pas aan op basis van gewichtsverloop en voedingstoestand van uw hond.';

  showResult('db');
}

/* 2. DOG – BARF */
function calcDogBarf() {
  const kg   = val('dbarf-weight');
  const life = radio('dbarf-life');
  const act  = radio('dbarf-act');
  const bcs  = radio('dbarf-bcs');
  const cast = checked('dbarf-castrated');

  if (!kg || kg < 0.5) { alert('Vul een geldig gewicht in.'); return; }

  const pct   = barfPctDog(life, act, bcs, cast);
  const total = (pct / 100) * kg * 1000; // g/day

  // Standard BARF composition for dogs
  g('dbarf-total').textContent   = Math.round(total);
  g('dbarf-pct-out').textContent = `${pct}% van lichaamsgewicht`;
  g('dbarf-meat').textContent    = fmt(total * 0.70);
  g('dbarf-bone').textContent    = fmt(total * 0.10);
  g('dbarf-organ').textContent   = fmt(total * 0.10);
  g('dbarf-veg').textContent     = fmt(total * 0.07);
  g('dbarf-suppl').textContent   = fmt(total * 0.03);

  const notes = [];
  if (life === 'puppy') notes.push('Pups: verhoog botaandeel tijdelijk tot maximaal 15% voor optimale calciumopname. Start met 5% en bouw langzaam op.');
  if (bcs === 'over')   notes.push('Overgewicht: verlaag tijdelijk naar 1,5% BW/dag en verhoog beweging.');
  g('dbarf-note').textContent = notes.join(' ') || 'Geef de dagportie in 2 maaltijden. Varieer eiwitbronnen (kip, rund, lam, vis) wekelijks voor een gevarieerd micronutriëntenprofiel.';

  showResult('dbarf');
}

/* 3. DOG – VERS GEKOOKT */

// kcal per 100 g rauw vlees (USDA FoodData Central gemiddeldes)
const PROTEIN_KCAL = { chicken:165, beef:250, pork:242, fish:208, turkey:189, lamb:294 };
// [proteinFraction, vegFraction, carbFraction]
const CARB_SPLIT   = { none:[0.70, 0.30, 0.00], low:[0.50, 0.25, 0.25], mid:[0.40, 0.25, 0.35] };

function calcDogCook() {
  const kg   = val('dcook-weight');
  const life = radio('dcook-life');
  const act  = radio('dcook-act');
  const cast = checked('dcook-castrated');
  const prot = g('dcook-protein').value;
  const carb = radio('dcook-carb');

  if (!kg || kg < 0.5) { alert('Vul een geldig gewicht in.'); return; }

  const RER    = rer(kg);
  const factor = dogMerFactor(life, act, 'ideal', cast, false);
  const MER    = RER * factor;

  const [protPct, vegPct, carbPct] = CARB_SPLIT[carb] || CARB_SPLIT.low;

  // Weighted kcal density of mixed meal (weighted average)
  const protKcal  = PROTEIN_KCAL[prot] || 165;
  const vegKcal   = 30;   // cooked vegetables avg kcal/100g
  const carbKcal  = 130;  // cooked rice / zoete aardappel avg kcal/100g
  const mixedDens = protPct * protKcal + vegPct * vegKcal + carbPct * carbKcal;
  const totalG    = (MER / mixedDens) * 100;
  const meals     = mealsPerDay(life);

  g('dcook-total').textContent    = Math.round(totalG);
  g('dcook-kcal-out').textContent = `≈ ${Math.round(MER)} kcal / dag`;
  g('dcook-mer').textContent      = Math.round(MER) + ' kcal/dag';
  g('dcook-meals').textContent    = meals + ' maaltijden';
  g('dcook-prot-pct').textContent = Math.round(protPct * 100) + '%';
  g('dcook-prot-g').textContent   = fmt(totalG * protPct);
  g('dcook-veg-g').textContent    = fmt(totalG * vegPct);
  g('dcook-carb-pct').textContent = Math.round(carbPct * 100) + '%';
  g('dcook-carb-g').textContent   = fmt(totalG * carbPct);

  g('dcook-note').textContent = 'Vers bereid voer bevat meer vocht dan brokken; de portie ziet er groter uit. Voeg altijd een complete hondenmultivitamine toe die geschikt is voor zelfbereide voeding.';

  showResult('dcook');
}

/* 4. CAT – DROOGVOER */
function calcCatDry() {
  const kg         = val('cd-weight');
  const kcalPer100 = val('cd-kcal');
  const life       = radio('cd-life');
  const style      = radio('cd-style');
  const bcs        = radio('cd-bcs');
  const cast       = checked('cd-castrated');
  const preg       = checked('cd-pregnant');

  if (!kg || kg < 0.3)                 { alert('Vul een geldig gewicht in.'); return; }
  if (!kcalPer100 || kcalPer100 < 100) { alert('Vul een geldige caloriewaarde in.'); return; }

  const RER    = rer(kg);
  const factor = catMerFactor(life, style, bcs, cast, preg);
  const MER    = RER * factor;
  const grams  = (MER / kcalPer100) * 100;
  const meals  = mealsPerDay(life);

  g('cd-grams').textContent    = Math.round(grams);
  g('cd-kcal-out').textContent = `≈ ${Math.round(MER)} kcal / dag`;
  g('cd-meals').textContent    = meals + ' maaltijden (of vrij eten)';
  g('cd-per-meal').textContent = Math.round(grams / meals) + ' g per maaltijd';
  g('cd-mer').textContent      = Math.round(MER) + ' kcal/dag';
  g('cd-factor').textContent   = factor + '× RER';

  const notes = [];
  if (preg)            notes.push('Zoogende poezen hebben tot 3× hogere energiebehoefte; laat voer altijd beschikbaar (vrij eten).');
  if (life === 'kitten') notes.push('Kittens: vrij eten aanbevolen tot ca. 6 maanden leeftijd.');
  if (bcs === 'over')    notes.push('Overgewicht: verlaag portie met 10–20% en controleer wekelijks het gewicht.');
  g('cd-note').textContent = notes.join(' ') || 'Katten drinken minder snel water; zorg naast droogvoer altijd voor vers (lopen) water of overweeg gedeeltelijk natvoer.';

  showResult('cd');
}

/* 5. CAT – BARF */
function calcCatBarf() {
  const kg    = val('cbarf-weight');
  const life  = radio('cbarf-life');
  const style = radio('cbarf-style');
  const bcs   = radio('cbarf-bcs');
  const cast  = checked('cbarf-castrated');

  if (!kg || kg < 0.3) { alert('Vul een geldig gewicht in.'); return; }

  const pct   = barfPctCat(life, style, bcs, cast);
  const total = (pct / 100) * kg * 1000;

  // Cat BARF: higher organ, NO vegetables (cats are strict carnivores)
  g('cbarf-total').textContent   = Math.round(total);
  g('cbarf-pct-out').textContent = `${pct}% van lichaamsgewicht`;
  g('cbarf-meat').textContent    = fmt(total * 0.65);
  g('cbarf-bone').textContent    = fmt(total * 0.12);
  g('cbarf-liver').textContent   = fmt(total * 0.10);
  g('cbarf-organ').textContent   = fmt(total * 0.10);
  g('cbarf-suppl').textContent   = fmt(total * 0.03);

  const notes = ['Voeg taurine toe (min. 200 mg/dag) — absoluut essentieel voor katten en niet aanwezig in plantaardige bronnen.'];
  if (life === 'kitten') notes.push('Kittens: verhoog botaandeel tijdelijk naar 15% voor optimale botopbouw.');
  g('cbarf-note').textContent = notes.join(' ');

  showResult('cbarf');
}

/* ════════════════════════════════════════════════════════════
   UI NAVIGATION
════════════════════════════════════════════════════════════ */

// Species switch (Hond / Kat)
document.querySelectorAll('.species-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.species-btn').forEach(b => { b.classList.remove('active'); b.setAttribute('aria-selected','false'); });
    btn.classList.add('active');
    btn.setAttribute('aria-selected','true');
    document.querySelectorAll('.species-panel').forEach(p => p.classList.remove('active'));
    g('panel-' + btn.dataset.species).classList.add('active');
  });
});

// Calculator type tabs (per species)
document.querySelectorAll('.ctype-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const tabBar      = btn.closest('.calc-type-tabs');
    const specPanel   = btn.closest('.species-panel');
    tabBar.querySelectorAll('.ctype-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    specPanel.querySelectorAll('.calc-panel').forEach(p => p.classList.remove('active'));
    g(btn.dataset.panel).classList.add('active');
  });
});