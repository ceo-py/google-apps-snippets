var workingSheet = "CharacterSettings";
var dataSheetName = "DataItems";
var ActiveColumns = [2, 9, 5, 12, 17, 22, 35];
var ActiveRows = Array.from({ length: 17 }, (_, index) => index + 7);
ActiveRows.push(4);

var gameVersionAllowItems = {
  sheetName: "ItemDependency",
  indexes: {
    0: "Head",
    1: "Chest",
    2: "Legs",
    3: "Shoulders",
    4: "Hands",
    5: "Wrists",
    6: "Waist",
    7: "Feet",
    8: "Back",
    9: "Neck",
    10: "Finger",
    11: "Finger",
    12: "Trinket",
    13: "Trinket",
    14: "Main Hand",
    15: "Off Hand",
    16: "Libram",
  },
  Head: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Conqueror's Aegis Headpiece (232)",
      "Ancient Iron Heaume (232)",
      "Steamworker's Goggles (232)",
      "Unwavering Stare (232)",
      "Collar of the Wyrmhunter (232)",
      "Valorous Aegis Headpiece (225)",
      "Faceguard of the Eyeless Horror (225)",
      "Helm of Veiled Energies (225)",
      "Cover of the Keepers (225)",
      "Lifespark Visage (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Conqueror's Aegis Headpiece (226)",
      "Ancient Iron Heaume (226)",
      "Steamworker's Goggles (226)",
      "Unwavering Stare (226)",
      "Collar of the Wyrmhunter (226)",
      "Valorous Aegis Headpiece (219)",
      "Faceguard of the Eyeless Horror (219)",
      "Helm of Veiled Energies (219)",
      "Cover of the Keepers (219)",
      "Lifespark Visage (219)",
    ],
  },
  Shoulders: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Observer's Mantle (239)",
      "Amice of Inconceivable Horror (239)",
      "Conqueror's Aegis Spaulders (232)",
      "Razorscale Shoulderguards (232)",
      "Amice of the Stoic Watch (232)",
      "Malleable Steelweave Mantle (232)",
      "Mantle of Wavering Calm (232)",
      "Valorous Aegis Spaulders (225)",
      "Pauldrons of Tempered Will (225)",
      "Pauldrons of the Combatant (252)",
      "Shoulderguards of Assimilation (225)",
      "Underworld Mantle (225)",
    ],

    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Pauldrons of the Combatant (239)",
      "Conqueror's Aegis Spaulders (226)",
      "Razorscale Shoulderguards (226)",
      "Observer's Mantle (226)",
      "Amice of the Stoic Watch (226)",
      "Malleable Steelweave Mantle (226)",
      "Mantle of Wavering Calm (226)",
      "Amice of Inconceivable Horror (226)",
      "Valorous Aegis Spaulders (219)",
      "Pauldrons of Tempered Will (219)",
      "Shoulderguards of Assimilation (219)",
      "Underworld Mantle (219)",
    ],
  },
  Back: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Drape of the Sullen Goddess (252)",
      "Sunglimmer Cloak (252)",
      "Sunglimmer Drape (239)",
      "Shroud of Alteration (232)",
      "Shawl of the Caretaker (225)",
      "Drape of the Spellweaver (225)",
      "Cloak of the Dormant Blaze (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Drape of the Sullen Goddess (239)",
      "Sunglimmer Cloak (239)",
      "Sunglimmer Drape (226)",
      "Shroud of Alteration (226)",
      "Shawl of the Caretaker (219)",
      "Drape of the Spellweaver (219)",
      "Cloak of the Dormant Blaze (219)",
    ],
  },
  Trinket: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Show of Faith (252)",
      "Meteorite Crystal (239)",
      "Sif's Remembrance (239)",
      "Pandora's Plea (232)",
      "Scale of Fates (232)",
      "Eye of the Broodmother (225)",
      "Spark of Hope (225)",
      "Energy Siphon (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Show of Faith (239)",
      "Meteorite Crystal (226)",
      "Sif's Remembrance (226)",
      "Pandora's Plea (226)",
      "Scale of Fates (226)",
      "Eye of the Broodmother (219)",
      "Spark of Hope (219)",
      "Energy Siphon (219)",
    ],
  },
  Chest: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Breastplate of the Devoted (252)",
      "Vestments of the Blind Denizen (252)",
      "Breastplate of the Stoneshaper (239)",
      "Conqueror's Aegis Tunic (232)",
      "Lifeforge Breastplate (232)",
      "Quartz-studded Harness (232)",
      "Chestguard of the Fallen God (232)",
      "Raiments of the Iron Council (232)",
      "Robes of the Umbral Brute (232)",
      "Valorous Aegis Tunic (225)",
      "Breastplate of the Afterlife (225)",
      "Firestrider Chestguard (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Breastplate of the Devoted (239)",
      "Vestments of the Blind Denizen (239)",
      "Conqueror's Aegis Tunic (226)",
      "Lifeforge Breastplate (226)",
      "Breastplate of the Stoneshaper (226)",
      "Quartz-studded Harness (226)",
      "Chestguard of the Fallen God (226)",
      "Raiments of the Iron Council (226)",
      "Robes of the Umbral Brute (226)",
      "Valorous Aegis Tunic (219)",
      "Breastplate of the Afterlife (219)",
      "Firestrider Chestguard (219)",
    ],
  },
  Hands: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Pharos Gloves (252)",
      "Gauntlets of the Thunder God (239)",
      "Gloves of Whispering Winds (239)",
      "Pulsar Gloves (239)",
      "Conqueror's Aegis Gloves (232)",
      "Gloves of the Pythonic Guardian (232)",
      "Gauntlets of Serene Blessing (232)",
      "Runeshaper's Gloves (232)",
      "Gloves of Augury (232)",
      "Grips of the Unbroken (232)",
      "Handwraps of Plentiful Recovery (232)",
      "Touch of the Occult (232)",
      "Valorous Aegis Gloves (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Pharos Gloves (239)",
      "Conqueror's Aegis Gloves (226)",
      "Gauntlets of the Thunder God (226)",
      "Gloves of the Pythonic Guardian (226)",
      "Gauntlets of Serene Blessing (226)",
      "Runeshaper's Gloves (226)",
      "Gloves of Augury (226)",
      "Gloves of Whispering Winds (226)",
      "Grips of the Unbroken (226)",
      "Handwraps of Plentiful Recovery (226)",
      "Touch of the Occult (226)",
      "Pulsar Gloves (226)",
      "Valorous Aegis Gloves (219)",
    ],
  },
  Finger: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "(Classic) Etched Signet of the Kirin Tor (232)",
      "Starshine Circle (252)",
      "Conductive Seal (252)",
      "Starshine Signet (239)",
      "Nebula Band (239)",
      "Fire Orchid Signet (239)",
      "Pyrelight Circle (232)",
      "Radiant Seal (232)",
      "Ring of the Faithful Servant (232)",
      "Frozen Loop (232)",
      "Sanity's Bond (232)",
      "Lady Maye's Sapphire Ring (225)",
      "Inscribed Signet of the Kirin Tor (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "(OG) Etched Signet of the Kirin Tor (226)",
      "Starshine Circle (239)",
      "Conductive Seal (239)",
      "Starshine Signet (226)",
      "Nebula Band (226)",
      "Fire Orchid Signet (226)",
      "Pyrelight Circle (226)",
      "Radiant Seal (226)",
      "Ring of the Faithful Servant (226)",
      "Frozen Loop (226)",
      "Sanity's Bond (226)",
      "Lady Maye's Sapphire Ring (219)",
      "Inscribed Signet of the Kirin Tor (213)",
    ],
  },
  Legs: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Zodiac Leggings (239)",
      "Conqueror's Aegis Greaves (232)",
      "Frostplate Greaves (232)",
      "Legguards of the Peaceful Covenant (232)",
      "Leggings of the Stoneweaver (232)",
      "Leggings of the Tortured Earth (232)",
      "Leggings of the Weary Mystic (232)",
      "Leggings of the Lifetender (232)",
      "Overload Legwraps (232)",
      "Leggings of Lost Love (232)",
      "Valorous Aegis Greaves (225)",
      "Legplates of Flourishing Resolve (225)",
      "Ironscale Leggings (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Conqueror's Aegis Greaves (226)",
      "Frostplate Greaves (226)",
      "Legguards of the Peaceful Covenant (226)",
      "Leggings of the Stoneweaver (226)",
      "Leggings of the Tortured Earth (226)",
      "Leggings of the Weary Mystic (226)",
      "Zodiac Leggings (226)",
      "Leggings of the Lifetender (226)",
      "Overload Legwraps (226)",
      "Leggings of Lost Love (226)",
      "Valorous Aegis Greaves (219)",
      "Legplates of Flourishing Resolve (219)",
      "Ironscale Leggings (219)",
    ],
  },
  Wrists: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Bindings of Winter Gale (252)",
      "Grasps of Reason (252)",
      "Horologist's Wristguards (232)",
      "Unfaltering Armguards (232)",
      "Wristguards of the Firetender (232)",
      "Bracers of the Broodmother (232)",
      "Bracers of Unleashed Magic (232)",
      "Armbands of the Construct (225)",
      "Bracers of Righteous Reformation (225)",
      "Armbraces of the Vibrant Flame (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Bindings of Winter Gale (239)",
      "Grasps of Reason (239)",
      "Horologist's Wristguards (226)",
      "Unfaltering Armguards (226)",
      "Wristguards of the Firetender (226)",
      "Bracers of the Broodmother (226)",
      "Bracers of Unleashed Magic (226)",
      "Armbands of the Construct (219)",
      "Bracers of Righteous Reformation (219)",
      "Armbraces of the Vibrant Flame (219)",
    ],
  },
  Neck: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Charm of Meticulous Timing (252)",
      "Sapphire Amulet of Renewal (252)",
      "Pendant of Fiery Havoc (252)",
      "Watchful Eye of Fate (239)",
      "Pendant of the Shallow Grave (239)",
      "Pendant of the Somber Witness (239)",
      "Freya's Choker of Warding (232)",
      "Evoker's Charm (232)",
      "Frozen Tear of Elune (232)",
      "Pendant of the Piercing Glare (225)",
      "Pendant of Endless Despair (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Charm of Meticulous Timing (239)",
      "Sapphire Amulet of Renewal (239)",
      "Pendant of Fiery Havoc (239)",
      "Watchful Eye of Fate (226)",
      "Pendant of the Shallow Grave (226)",
      "Pendant of the Somber Witness (226)",
      "Freya's Choker of Warding (226)",
      "Evoker's Charm (226)",
      "Frozen Tear of Elune (226)",
      "Pendant of the Piercing Glare (219)",
      "Pendant of Endless Despair (219)",
    ],
  },
  "Main Hand": {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Val'anyr, Hammer of Ancient Kings (258)",
      "Constellus (252)",
      "Aesuga, Hand of the Ardent Champion (245)",
      "Fusion Blade (245)",
      "Guiding Star (238)",
      "Firesoul (225)",
      "Pulse Baton (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Val'anyr, Hammer of Ancient Kings (245)",
      "Constellus (239)",
      "Aesuga, Hand of the Ardent Champion (232)",
      "Fusion Blade (232)",
      "Guiding Star (232)",
      "Firesoul (219)",
      "Pulse Baton (219)",
    ],
  },
  Waist: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Star-beaded Clutch (252)",
      "Belt of the Crystal Tree (239)",
      "Belt of Clinging Hope (232)",
      "Plate Girdle of Righteousness (232)",
      "Girdle of Unyielding Trust (232)",
      "Belt of the Fallen Wyrm (232)",
      "Blue Belt of Chaos (232)",
      "Windchill Binding (232)",
      "Belt of the Sleeper (232)",
      "Belt of Arctic Life (232)",
      "Belt of the Living Thicket (232)",
      "Embrace of the Leviathan (232)",
      "Sash of Ancient Power (232)",
      "Cable of the Metrognome (225)",
      "Belt of the Iron Servant (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Star-beaded Clutch (239)",
      "Belt of Clinging Hope (226)",
      "Plate Girdle of Righteousness (226)",
      "Girdle of Unyielding Trust (226)",
      "Belt of the Fallen Wyrm (226)",
      "Blue Belt of Chaos (226)",
      "Windchill Binding (226)",
      "Belt of the Crystal Tree (226)",
      "Belt of the Sleeper (226)",
      "Belt of Arctic Life (226)",
      "Belt of the Living Thicket (226)",
      "Embrace of the Leviathan (226)",
      "Sash of Ancient Power (226)",
      "Cable of the Metrognome (219)",
      "Belt of the Iron Servant (219)",
    ],
  },
  Feet: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Planewalker Treads (252)",
      "Treads of the False Oracle (252)",
      "Boots of Fiery Resolution (252)",
      "Greaves of the Rockmender (232)",
      "Treads of Destiny (232)",
      "Boots of the Forgotten Depths (232)",
      "Lightning Grounded Boots (232)",
      "Boots of the Servant (232)",
      "Spellslinger's Slippers (232)",
      "Sabatons of the Iron Watcher (225)",
      "Greaves of the Earthbinder (225)",
      "Boots of the Petrified Forest (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Planewalker Treads (239)",
      "Boots of Fiery Resolution (239)",
      "Treads of the False Oracle (239)",
      "Greaves of the Rockmender (226)",
      "Treads of Destiny (226)",
      "Boots of the Forgotten Depths (226)",
      "Lightning Grounded Boots (226)",
      "Boots of the Servant (226)",
      "Spellslinger's Slippers (226)",
      "Sabatons of the Iron Watcher (219)",
      "Greaves of the Earthbinder (219)",
      "Boots of the Petrified Forest (219)",
    ],
  },
  "Off Hand": {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Wisdom's Hold (252)",
      "Ice Layered Barrier (245)",
      "Pulsing Spellshield (225)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Wisdom's Hold (239)",
      "Ice Layered Barrier (232)",
      "Pulsing Spellshield (219)",
    ],
  },
  Libram: {
    "Wrath OG (v3.3.5)": [
      "Phase 2 (Classic)",
      "Libram of the Resolute (232)",
      "Furious Gladiator's Libram of Justice (238)",
    ],
    "Wrath Classic (v3.4.3)": [
      "Phase 2 (OG)",
      "Libram of the Resolute (226)",
      "Furious Gladiator's Libram of Justice (232)",
    ],
  },
};

var raceAndGameVersion = { row: 4, column: [2, 9] };
var gameVersions = ["Wrath OG (v3.3.5)", "Wrath Classic (v3.4.3)"];
var races = {
  horde: ["Blood Elf"],
  aliance: ["Human", "Dwarf", "Draenei"],
};

var skipingUniqueItems = [
  "(A) Solace of the Defeated (245)",
  "(A) Solace of the Defeated (258)",
  "(A) Ring of the Darkmender (245)",
  "(A) Ring of the Darkmender (258)",
  "(H) Solace of the Fallen (245)",
  "(H) Solace of the Fallen (258)",
  "(H) Circle of the Darkmender (245)",
  "(H) Circle of the Darkmender (258)",
];
var enchantHeadOptionsDropDown = [
  "None",
  "Arcanum of Burning Mysteries",
  "Arcanum of Blissful Mending",
];
var enchantShoulderOptionsDropDown = [
  "None",
  "Master's Inscription of the Storm",
  "Master's Inscription of the Crag",
  "Greater Inscription of the Storm",
  "Greater Inscription of the Crag",
];
var enchantBackOptionsDropDown = [
  "None",
  "Springy Arachnoweave",
  "Lightweave Embroidery ",
  "Darkglow Embroidery ",
  "Greater Speed",
];
var enchantChestOptionsDropDown = [
  "None",
  "Powerful Stats",
  "Greater Mana Restoration",
  "Exceptional Mana",
];
var enchantHandsOptionsDropDown = [
  "None",
  "Hyperspeed Accelerators",
  "Exceptional Spellpower",
  "Minor Haste",
];
var enchantWristOptionsDropDown = [
  "None",
  "Fur Lining - Spell Power",
  "Superior Spellpower",
  "Exceptional Intellect",
];
var enchantsLegsOptionsDropDown = [
  "None",
  "Brilliant Spellthread",
  "Sapphire Spellthread",
];
var enchantsFeetOptionsDropDown = [
  "None",
  "Nitro Boosts",
  "Icewalker",
  "Greater Vitality",
  "Tuskarr's Vitality",
];
var enchantFingerOptionsDropDown = ["None", "Greater Spellpower"];
var enchantMainHandOptionsDropDown = [
  "None",
  "Mighty Spellpower",
  "Major Intellect",
];
var enchantOffHandOptionsDropDown = ["None", "Greater Intellect"];
var enchantOptions = {
  7: enchantHeadOptionsDropDown,
  8: enchantChestOptionsDropDown,
  9: enchantsLegsOptionsDropDown,
  10: enchantShoulderOptionsDropDown,
  11: enchantHandsOptionsDropDown,
  12: enchantWristOptionsDropDown,
  14: enchantsFeetOptionsDropDown,
  15: enchantBackOptionsDropDown,
  17: enchantFingerOptionsDropDown,
  18: enchantFingerOptionsDropDown,
  21: enchantMainHandOptionsDropDown,
  22: enchantOffHandOptionsDropDown,
};
var excludeGemNames = [
  "Prismatic",
  "Yellow",
  "Red",
  "Orange",
  "Purple",
  "Green",
];
var colors = {
  0: "L",
  1: "Q",
  2: "V",
};
var backgroundColors = {
  Meta: "#bdbdbd",
  Blue: "#0096ff",
  Red: "Red",
  Yellow: "Yellow",
};
var backgroundColorsBonus = {
  "#bdbdbd": ["#bdbdbd"],
  "#0096ff": [
    "Nightmare Tear",
    "Enchanted Tear",
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
    "Royal Dreadstone",
    "Royal Twilight Opal",
    "Perfect Royal Shadow Crystal",
    "Royal Shadow Crystal",
  ],
  Red: [
    "Nightmare Tear",
    "Enchanted Tear",
    "Luminous Ametrine",
    "Reckless Ametrine",
    "Luminous Monarch Topaz",
    "Reckless Monarch Topaz",
    "Perfect Luminous Huge Citrine",
    "Perfect Reckless Huge Citrine",
    "Luminous Huge Citrine",
    "Reckless Huge Citrine",
    "Royal Dreadstone",
    "Royal Twilight Opal",
    "Perfect Royal Shadow Crystal",
    "Royal Shadow Crystal",
    "Runed Dragon's Eye",
    "Runed Cardinal Ruby",
    "Runed Scarlet Ruby",
    "Perfect Runed Bloodstone",
    "Runed Bloodstone",
  ],
  Yellow: [
    "Nightmare Tear",
    "Enchanted Tear",
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
    "Luminous Ametrine",
    "Reckless Ametrine",
    "Luminous Monarch Topaz",
    "Reckless Monarch Topaz",
    "Perfect Luminous Huge Citrine",
    "Perfect Reckless Huge Citrine",
    "Luminous Huge Citrine",
    "Reckless Huge Citrine",
    "Brilliant Dragon's Eye",
    "Quick Dragon's Eye",
    "Brilliant King's Amber",
    "Quick King's Amber",
    "Brilliant Autumn's Glow",
    "Quick Autumn's Glow",
    "Perfect Brilliant Sun Crystal",
    "Perfect Quick Sun Crystal",
    "Brilliant Sun Crystal",
    "Quick Sun Crystal",
  ],
};
var bsAndSocketBelt = {
  1135: "AI11",
  1235: "AI12",
  1335: "AI13",
  115: "AI11",
  125: "AI12",
  135: "AI13",

  1112: "AI11",
  1117: "AI11",
  1122: "AI11",

  1212: "AI12",
  1217: "AI12",
  1222: "AI12",

  1312: "AI13",
  1317: "AI13",
  1322: "AI13",
};
var correctGemColors = {
  Yellow: [
    "Brilliant Dragon's Eye",
    "Quick Dragon's Eye",
    "Brilliant King's Amber",
    "Quick King's Amber",
    "Brilliant Autumn's Glow",
    "Quick Autumn's Glow",
    "Perfect Brilliant Sun Crystal",
    "Perfect Quick Sun Crystal",
    "Brilliant Sun Crystal",
    "Quick Sun Crystal",
  ],
  Red: [
    "Runed Dragon's Eye",
    "Runed Cardinal Ruby",
    "Runed Scarlet Ruby",
    "Perfect Runed Bloodstone",
    "Runed Bloodstone",
  ],
  Orange: [
    "Luminous Ametrine",
    "Reckless Ametrine",
    "Luminous Monarch Topaz",
    "Reckless Monarch Topaz",
    "Perfect Luminous Huge Citrine",
    "Perfect Reckless Huge Citrine",
    "Luminous Huge Citrine",
    "Reckless Huge Citrine",
  ],
  Purple: [
    "Royal Dreadstone",
    "Royal Twilight Opal",
    "Perfect Royal Shadow Crystal",
    "Royal Shadow Crystal",
  ],
  Green: [
    "Dazzling Eye of Zul",
    "Energized Eye of Zul",
    "Dazzling Forest Emerald",
    "Energized Forest Emerald",
    "Perfect Dazzling Dark Jade",
    "Perfect Energized Dark Jade",
    "Dazzling Dark Jade",
    "Energized Dark Jade",
  ],
};
var correctNumberGems = {
  Tear: 1,
  "Dragon's": 3,
};
var dropdownOptions = [
  "Prismatic",
  "Nightmare Tear",
  "Enchanted Tear",
  "Yellow",
  "Brilliant Dragon's Eye",
  "Quick Dragon's Eye",
  "Brilliant King's Amber",
  "Quick King's Amber",
  "Brilliant Autumn's Glow",
  "Quick Autumn's Glow",
  "Perfect Brilliant Sun Crystal",
  "Perfect Quick Sun Crystal",
  "Brilliant Sun Crystal",
  "Quick Sun Crystal",
  "Red",
  "Runed Dragon's Eye",
  "Runed Cardinal Ruby",
  "Runed Scarlet Ruby",
  "Perfect Runed Bloodstone",
  "Runed Bloodstone",
  "Orange",
  "Luminous Ametrine",
  "Reckless Ametrine",
  "Luminous Monarch Topaz",
  "Reckless Monarch Topaz",
  "Perfect Luminous Huge Citrine",
  "Perfect Reckless Huge Citrine",
  "Luminous Huge Citrine",
  "Reckless Huge Citrine",
  "Purple",
  "Royal Dreadstone",
  "Royal Twilight Opal",
  "Perfect Royal Shadow Crystal",
  "Royal Shadow Crystal",
  "Green",
  "Dazzling Eye of Zul",
  "Energized Eye of Zul",
  "Dazzling Forest Emerald",
  "Energized Forest Emerald",
  "Perfect Dazzling Dark Jade",
  "Perfect Energized Dark Jade",
  "Dazzling Dark Jade",
  "Energized Dark Jade",
];
var dropDownOptionsMeta = [
  "Meta",
  "Insightful Earthsiege Diamond",
  "Ember Skyflare Diamond",
  "Beaming Earthsiege Diamond",
  "Revitalizing Skyflare Diamond",
];
var metaActivationReq = {
  "Insightful Earthsiege Diamond": { Red: 1, Yellow: 1, Blue: 1 },
  "Ember Skyflare Diamond": { Red: 3 },
  "Beaming Earthsiege Diamond": { Red: 2, Yellow: 1 },
  "Revitalizing Skyflare Diamond": { Red: 2 },
};
var setBonus = {
  range: { row: [7, 8, 9, 10, 11], col: 5 },
  bonusActiveRange: ["P26", "P27", "P28", "P29", "P30", "P31", "P32", "P33"],
  tiers: {
    tier7: {
      name: "Heroes' Redemption",
      "2 Pieces": "P26",
      "4 Pieces": "P27",
    },
    tier71: {
      name: "Valorous Redemption",
      "2 Pieces": "P26",
      "4 Pieces": "P27",
    },
    tier8: {
      name: "Valorous Aegis",
      "2 Pieces": "P28",
      "4 Pieces": "P29",
    },
    tier81: {
      name: "Conqueror's Aegis",
      "2 Pieces": "P28",
      "4 Pieces": "P29",
    },
    tier9: {
      name: "Turalyon's",
      "2 Pieces": "P30",
      "4 Pieces": "P31",
    },
    tier91: {
      name: "Liadrin's",
      "2 Pieces": "P30",
      "4 Pieces": "P31",
    },
    tier10: {
      name: "Lightsworn",
      "2 Pieces": "P32",
      "4 Pieces": "P33",
    },
  },
};

function applyColorBaseOnItemSockets(e) {
  const spreadSheet = e.source;
  const activeRange = e.source.getActiveRange();
  const sheetName = spreadSheet.getActiveSheet().getName();

  if (sheetName !== workingSheet) return;

  const activeSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const row = activeRange.getRow();
  const column = activeRange.getColumn();

  metaNotActive(activeSheet);

  if (!ActiveColumns.includes(column) || !ActiveRows.includes(row)) return;
  const rowInDataSheet = activeSheet.getRange(`A${row}`).getValue();
  const currentCellValue = activeRange.getValue();

  correctHordAlianceItems(activeSheet, currentCellValue);

  if (gameVersions.includes(currentCellValue)) {
    itemsCorrectGameVersion(activeSheet, currentCellValue);
  }

  if (
    !(column === 35 && [11, 12, 13].includes(row)) &&
    containsExactWord(currentCellValue, "Phase")
  ) {
    resetGemCells(row, activeSheet);
    isSetBonus(activeSheet);
    return;
  }
  // activeSheet.getRange("H5").setValue(`${row} ${column}`);

  const additionalSocket = bsAndSocketBelt?.[`${row}${column}`];
  const dataSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
  const targetRange = dataSheet.getRange(
    `N${rowInDataSheet}:P${rowInDataSheet}`
  );
  const data = targetRange.getValues()[0].filter((gem) => gem !== "");
  // activeSheet.getRange("I5").setValue(additionalSocket);

  if ([12, 17, 22, 35].includes(column)) {
    isGemsCorrectPrismaticAndDragons(
      activeSheet,
      currentCellValue,
      activeRange
    );
  }

  const gameVersion = activeSheet.getRange("B4").getValue();

  if (
    column === 5 &&
    [17, 18, 19, 20].includes(row) &&
    gameVersion === "Wrath Classic (v3.4.3)" &&
    !isItemUnique(activeSheet, activeRange, row)
  )
    return;
  else if (
    column === 5 &&
    gameVersion === "Wrath OG (v3.3.5)" &&
    [17, 18, 19, 20].includes(row) &&
    !isRingTrinketUniqueOldVersion(
      activeSheet,
      activeRange,
      currentCellValue,
      row
    )
  ) {
    return;
  }
  if (![12, 17, 22, 35].includes(column)) resetGemCells(row, activeSheet);
  // Extra Socket Change only
  if (
    column === 35 &&
    additionalSocket &&
    !activeSheet.getRange(additionalSocket).getValue()
  ) {
    resetOnlyExtraGemSlot(activeSheet, row, data.length);
  }

  if (
    additionalSocket &&
    activeSheet.getRange(additionalSocket).getValue() &&
    !containsExactWord(activeSheet.getRange(`E${row}`).getValue(), "Phase")
  ) {
    data.push("Meta");
  }

  var bonusForGems = generateGemsDropDownMenu(activeSheet, row, data);
  const isBonusSocket = bonusForGems === data.length && bonusForGems > 0;

  if ([5, 12, 17, 22].includes(column)) {
    setBonusSocket(activeSheet, row, isBonusSocket, data);
  }

  if ([5].includes(column)) {
    enchantsDropDown(activeSheet, row);
  }

  isMetaActive(activeSheet, activeSheet.getRange("L7").getValue());

  if (setBonus.range.row.includes(row) && setBonus.range.col === column) {
    isSetBonus(activeSheet);
  }

  // itemsCorrectGameVersion(activeSheet);

  // activeSheet
  //   .getRange("J5")
  //   .setValue(`${isBonusSocket} ${bonusForGems} ${data.length}`);
}
function itemsCorrectGameVersionUiqueCheck(activeSheet) {
  const uniqueItems = {
    "Solace of the": {
      resetRow: "19",
    },
    "of the Darkmender": {
      resetRow: "17",
    },
  };
  const ringsAndTrinkets = activeSheet
    .getRange("E17:E20")
    .getValues()
    .map((x) => x.toString());
  Object.keys(uniqueItems).forEach((key) => {
    if (
      ringsAndTrinkets.filter((x) => containsExactWord(x, key)).length === 2
    ) {
      resetGemCells(uniqueItems[key].resetRow, activeSheet);
      activeSheet.getRange(`E${uniqueItems[key].resetRow}`).setValue("");
    }
  });
}

function itemsCorrectGameVersion(activeSheet, gameVersion) {
  const items = activeSheet
    .getRange("E7:E23")
    .getValues()
    .map((x) => x.toString());
  items.forEach((item, i) => {
    if (
      gameVersionAllowItems[gameVersionAllowItems.indexes[i]][
        gameVersion
      ].includes(item)
    ) {
      resetGemCells(i + 7, activeSheet);
      activeSheet.getRange(`E${i + 7}`).setValue("");
    }
  });
  if (gameVersion === "Wrath Classic (v3.4.3)") {
    itemsCorrectGameVersionUiqueCheck(activeSheet);
  }
}

function correctHordAlianceItems(activeSheet, currentCellValue) {
  const items = activeSheet
    .getRange("E7:E23")
    .getValues()
    .map((x) => x.toString());

  items.forEach((item, i) => {
    if (
      (races.horde.includes(currentCellValue) && item.includes("(A)")) ||
      (races.aliance.includes(currentCellValue) && item.includes("(H)"))
    ) {
      resetGemCells(i + 7, activeSheet);
      activeSheet.getRange(`E${i + 7}`).setValue("");
    }
  });
}

function combineColors(initialDict, compareDict) {
  for (const color in compareDict) {
    if (compareDict.hasOwnProperty(color)) {
      const colorDict = compareDict[color];
      for (const key in colorDict) {
        if (colorDict.hasOwnProperty(key)) {
          if (initialDict.hasOwnProperty(key)) {
            initialDict[key] += colorDict[key];
          } else {
            initialDict[key] = colorDict[key];
          }
        }
      }
    }
  }
  return initialDict;
}
function getTotalGems(data) {
  const subColors = {
    Orange: { Red: 0, Yellow: 0 },
    Purple: { Red: 0, Blue: 0 },
    Green: { Yellow: 0, Blue: 0 },
  };
  const foundGems = {};
  for (const cValue of data) {
    if (cValue === "") continue;

    if (cValue === "Nightmare Tear" || cValue === "Enchanted Tear") {
      foundGems["Tear"] = 1;
      continue;
    }

    Object.keys(correctGemColors).forEach((c) => {
      const gemCorrect = correctGemColors[c].includes(cValue);
      if (subColors.hasOwnProperty(c) && gemCorrect) {
        for (const key in subColors[c]) {
          subColors[c][key] += 1;
        }
      } else if (gemCorrect) {
        if (!foundGems.hasOwnProperty(c)) foundGems[c] = 0;
        foundGems[c]++;
      }
    });
  }
  return combineColors(foundGems, subColors);
}

function tearGemAdd(metaType, foundGems) {
  if (foundGems.hasOwnProperty("Tear")) {
    for (const color in metaActivationReq[metaType]) {
      metaActivationReq[metaType][color]--;
      if (metaActivationReq[metaType][color] <= 0)
        delete metaActivationReq[metaType][color];
    }
    delete foundGems["Tear"];
  }
  return foundGems;
}

function getKeyWithMostValue(dict, topColor) {
  let maxKey = null;
  let maxValue = -Infinity;

  for (const key in dict) {
    let totalValue = 0;
    const colorValues = dict[key];
    for (const color in colorValues) {
      totalValue += colorValues[color];
    }
    if (totalValue > maxValue && dict[key].hasOwnProperty(topColor)) {
      maxValue = totalValue;
      maxKey = key;
    }
  }
  return maxKey;
}

function isColorsCorrect(initialDict, compareDict) {
  const initialSet = new Set(Object.keys(initialDict).map((x) => x));
  const compareSet = new Set(compareDict);
  for (let item of initialSet) {
    if (!compareSet.has(item)) {
      return false;
    }
  }
  return true;
}

function metaNotActive(activeSheet) {
  activeSheet.getRange("AI14").setValue("No");
  activeSheet.getRange("L7").setFontColor("Red");
}
function isMetaActive(sheet, metaType) {
  if (!metaActivationReq.hasOwnProperty(metaType)) return;

  const range = sheet.getRange("L7:V23");
  const values = range
    .getValues()
    .map((x) => JSON.stringify(x))
    .map((x) => JSON.parse(x))
    .filter((x) => x.length !== 0)
    .flat();

  let foundGems = getTotalGems(values);
  // sheet.getRange("I2").setValue(foundGems);

  foundGems = tearGemAdd(metaType, foundGems);

  for (const color in metaActivationReq[metaType]) {
    if (foundGems.hasOwnProperty(color)) {
      metaActivationReq[metaType][color] -= foundGems[color];
      if (metaActivationReq[metaType][color] <= 0)
        delete metaActivationReq[metaType][color];
      delete foundGems[color];
    }
  }

  const missingToActive = Object.values(metaActivationReq[metaType]).reduce(
    (a, b) => a + b,
    0
  );
  const metaActive = missingToActive === 0;

  if (metaActive) {
    sheet.getRange("L7").setFontColor("#6aa84f");
    sheet.getRange("AI14").setValue("Yes");
    return;
  }
  sheet.getRange("AI14").setValue("No");
  sheet.getRange("L7").setFontColor("Red");
}

function resetSetBonus(sheet) {
  setBonus.bonusActiveRange.forEach((cell) => {
    sheet.getRange(cell).setValue("No");
    sheet.getRange(cell).setFontColor("Red");
  });
}

function isSetBonus(sheet) {
  const setBonusItems = sheet.getRange("E7:E11").getValues();
  resetSetBonus(sheet);
  Object.keys(setBonus.tiers).forEach((key) => {
    const tier = setBonus.tiers[key];
    const countItems = setBonusItems.filter((x) =>
      containsExactWord(x, tier.name)
    ).length;
    const fourPeaceCell = sheet.getRange(tier["4 Pieces"]);
    const twoPeaceCell = sheet.getRange(tier["2 Pieces"]);
    // Browser.msgBox(`${tier.name} - ${countItems}`);
    if ([4, 5].includes(countItems)) {
      fourPeaceCell.setValue("Yes");
      fourPeaceCell.setFontColor("#6aa84f");
      twoPeaceCell.setValue("Yes");
      twoPeaceCell.setFontColor("#6aa84f");
    } else if ([2, 3].includes(countItems)) {
      twoPeaceCell.setValue("Yes");
      twoPeaceCell.setFontColor("#6aa84f");
    }
  });
}

function rangeRingsOrTrinkets(row) {
  const items = {
    ring: "E17:E18",
    trinket: "E19:E20",
  };

  const rings = [17, 18];
  if (rings.includes(row)) return items.ring;
  return items.trinket;
}

function isRingTrinketUniqueOldVersion(
  sheet,
  activeRange,
  currentCellValue,
  row
) {
  const items = sheet
    .getRange(rangeRingsOrTrinkets(row))
    .getValues()
    .flat()
    .map((x) => x.toString());

  if (
    items.length === new Set(items).size &&
    skipingUniqueItems.includes(currentCellValue)
  )
    return true;
  return isItemUnique(sheet, activeRange, row);
}

function isItemUnique(sheet, activeRange, row) {
  const range = sheet.getRange(rangeRingsOrTrinkets(row));
  const values = range
    .getValues()
    .flat()
    .map((x, i) =>
      x === "" ? i : containsExactWord(x, "Phase") ? i : x.split(" (")[0]
    );
  if (values.length !== new Set(values).size) {
    resetGemCells(row, sheet);
    activeRange.setValue("");
    return false;
  }
  return true;
}

function isGemsCorrectPrismaticAndDragons(sheet, activeCellValue, activeRange) {
  let gemType = "";
  for (const gem in correctNumberGems) {
    if (containsExactWord(activeCellValue, gem)) {
      gemType = gem;
      break;
    }
  }
  if (gemType === "") return;

  const range = sheet.getRange("L7:V23");
  const values = range
    .getValues()
    .map((x) => JSON.stringify(x).split(","))
    .flat();
  let found = 0;
  values.map((cellValue) => {
    if (containsExactWord(cellValue, gemType)) found++;
  });
  if (found > correctNumberGems[gemType]) activeRange.setValue("");
}

function containsExactWord(text, word) {
  const regex = new RegExp(`\\b${word}\\b`, "i");
  return regex.test(text);
}

function resetOnlyExtraGemSlot(sheet, row, gems) {
  const cellToChange = sheet.getRange(
    `${gems === 0 ? "L" : gems === 1 ? "Q" : gems === 2 ? "V" : ""}${row}`
  );
  cellToChange.setValue("").setDataValidation(null);
  cellToChange.setBackground("");
}

function resetGemCells(row, sheet) {
  Object.values(colors).forEach((cell) => {
    const cellToChange = sheet.getRange(`${cell}${row}`);
    sheet.getRange(`Z${row}`).setValue("");
    sheet.getRange(`AB${row}`).setValue("").setDataValidation(null);
    cellToChange.setDataValidation(null);
    cellToChange.setBackground("");
    cellToChange.setValue("");
  });
}

function enchantsDropDown(activeSheet, row) {
  const cellToChange = activeSheet.getRange(`AB${row}`);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(enchantOptions[row])
    .build();
  cellToChange.setDataValidation(rule);
}

function setBonusSocket(sheet, row, bonus, data) {
  const bonusCellToChange = sheet.getRange(`Z${row}`);
  bonusCellToChange.setValue(
    data.length === 0 || (data.length === 1 && data.includes("Meta"))
      ? ""
      : bonus
      ? "Yes"
      : "No"
  );
  bonusCellToChange.setFontColor(bonus ? "#6aa84f" : "Red");
}

function generateGemsDropDownMenu(activeSheet, row, data) {
  let bonusTotal = 0;
  data.forEach((color, i) => {
    const cellToChange = activeSheet.getRange(`${colors[i]}${row}`);
    const cellValue = cellToChange.getValue();
    // Browser.msgBox(color);
    bonusTotal += backgroundColorsBonus[backgroundColors[color]].includes(
      cellValue || ""
    )
      ? 1
      : 0;
    const isMetaColpr =
      (color === "Meta" && row !== 7 && cellValue !== "Meta") ||
      (row === 7 &&
        color === "Meta" &&
        cellValue !== "Meta" &&
        cellValue !== "");
    bonusTotal += isMetaColpr ? 1 : 0;
    cellToChange.setBackground(backgroundColors[color]);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        color === "Meta" && row === 7 ? dropDownOptionsMeta : dropdownOptions
      )
      .build();
    cellToChange.setDataValidation(rule);
  });
  return bonusTotal;
}
