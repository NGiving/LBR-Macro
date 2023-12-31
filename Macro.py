from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import mouse
from PIL import ImageGrab
from enum import Enum
import numpy as np
import cv2 as cv
import pytesseract
import win32api
from bisect import insort
from datetime import timedelta
from base64 import b64decode
import json
import math
import sched, time
from os import path, getenv
from sys import exit

class Rarity(str, Enum):
    COMMON = 'common'
    UNCOMMON = 'uncommon'
    RARE = 'rare'
    EPIC = 'epic'
    MYTHICAL = 'mythical'
    LEGENDARY = 'legendary'

class Artifact(str, Enum):
    LANTERN = 'lantern'
    LE = 'leafscension_exploit'
    ROD = 'nature_rod'
    FRUIT = 'fruit'
    GRAV = 'gravity_ball'
    ZOO = 'zoo'
    SEED = 'seed_bag'
    VORTEX = 'vortex'
    WINGS = 'wings'
    COMPASS = 'compass'
    SEAL = 'water_seal'
    SKULL = 'blazing_skull'
    TIME_CRYSTAL = 'time_crystal'
    SUITCASE = 'gold_suitcase'
    ORB = 'orb'
    VIOLIN = 'vital_violin'
    WIND = 'wind'

class Regular(str, Enum):
    ALB = 'alb'
    ANGRY_LEAF = 'angry_leaf'
    BAT = 'bat'
    BEE = 'bee'
    BIRD = 'bird'
    BUG = 'bug'
    CARROT = 'carrot'
    CRAB = 'crab'
    ELK = 'elk'
    FROG = 'frog'
    HARE = 'hare'
    HYENA = 'hyena'
    ICE_LEAF = 'ice_leaf'
    SCORPION = 'scorpion'
    SNOWMAN = 'snowman'
    SKULL = 'skull'
    VULTURE = 'vulture'
    TOWER_FROG = 'tower_frog'
    TOWER_BAT = 'evil_bat'
    TOWER_HARE = 'evil_hare'
    TOWER_CRAB = 'evil_crab'
    TOWER_SCORPION = 'evil_scorpion'
    TOWER_SNOWMAN = 'evil_snowman'
    TOWER_BRAIN = 'evil_brain'
    TOWER_BIRD =  'evil_bird'
    TOWER_CARROT = 'evil_carrot'
    TOWER_SKULL = 'evil_skull'
    TOWER_STONE = 'evil_stone'
    TOWER_FACTORY = 'evil_factory'
    TOWER_FIRE = 'evil_fire'
    TOWER_GHAST = 'evil_ghast'
    TOWER_HEAD = 'evil_head'
    TOWER_HULK = 'evil_hulk'
    TOWER_KLACKON = 'evil_klackon'
    TOWER_LURKER = 'evil_lurker'
    TOWER_MASK = 'evil_mask'
    TOWER_BAD_SKULL = 'evil_skull02'
    SOUL_MONOLITH = 'sc_monolith'
    SOUL_MUMMY = 'sc_mummy'
    SOUL_PUDDLE = 'sc_puddle'

class Boss(str, Enum):
    FROG = 'boss_frog'
    BAT = 'boss_bat'
    HARE = 'boss_hare'
    CRAB = 'boss_crab'
    SCORPION = 'boss_scorpion'
    SNOWMAN = 'boss_snowman'
    BRAIN = 'boss_brain'
    DRAGON = 'boss_dragon'
    LAVA_GOLEM = 'boss_lavagolem'
    STONE_GOLEM = 'boss_stonegolem'
    FACOTRY = 'boss_factory'
    FIRE = 'boss_fire'
    GHAST = 'boss_ghast'
    HEAD = 'boss_head'
    HULK = 'boss_hulk'
    KLACKON = 'boss_klackon'
    LURKER = 'boss_lurker'
    MASK = 'boss_mask'
    BOB_1 = 'boss_bob'
    BOB_2 = 'boss_bob02'
    BOB_3 = 'boss_bob03'
    BOB_4 = 'boss_bob04'
    BOB_5 = 'boss_bob05'
    WITCH = 'boss_witch'
    CENTAUR = 'boss_centaur'
    VILE = 'boss_vile_creature'
    ELEMENTAL = 'boss_air_elemental'
    BUBBLE = 'boss_spark_bubble'
    TERROR_BLUE = 'boss_terror_blue'
    TERROR_GREEN = 'boss_terror_green'
    TERROR_RED = 'boss_terror_red'
    TERROR_PURPLE = 'boss_terror_purple'
    TERROR_SUPER = 'boss_terror_super'
    ENERGY_GUARD = 'boss_energy_guard'
    GREEN_FLAME = 'boss_green_flame'
    SPECTRALSEEKER = 'boss_spectralseeker'
    MIRAGE = 'boss_soul_mirage'

class Area(str, Enum):
    GARDEN = 'garden'
    NEIGHBORS_GARDEN = 'neighbors_garden'
    MOUNTAIN = 'mountain'
    SPACE = 'space'
    VOID = 'void'
    ABYSS = 'abyss'
    CELESTIAL = 'celestial_plane'
    MYTHICAL_GARDEN = 'mythical_garden'
    VOLCANO = 'volcano'
    RESEARCH_STATION = 'antarctica'
    HIDDEN_SEA = 'hidden_sea'
    HARBOR = 'harbor'
    TOWER = 'tower'
    MOON = 'moon'
    DESERT = 'desert'
    PYRAMID = 'pyramid'
    INNER_PYRAMID = 'inner_pyramid'
    KOKKAUPUNKI = 'kokkaupunki'
    CURSED_KOKKAUPUNKI = 'kokkaupunki_cursed'
    GLADE = 'dark_glade'
    BLACK_HOLE = 'black_leaf_hole'
    DICEY_MEADOWS = 'dicey_meadows'
    GLINTING_THICKET = 'glinting_thicket'
    PUB = 'cheese_pub'
    HOUSE = 'house'
    BIOTITE = 'biotite_forest'
    BRIDGE = 'exalted_bridge'
    SANCTUM = 'ancient_sanctum'
    CEMETERY = 'vilewood_cemetery'
    LONE_TREE = 'lone_tree'
    RANGE = 'spark_range'
    BUBBLE = 'spark_bubble'
    SPARK_PORTAL = 'spark_portal'
    SHRINE = 'energy_shrine'
    PLASMA_FOREST = 'plasma_forest'
    BLUE_PLANET = 'planet_edge_blue'
    GREEN_PLANET = 'planet_edge_green'
    RED_PLANET = 'planet_edge_red'
    PURPLE_PLANET = 'planet_edge_purple'
    BLACK_PLANET = 'planet_edge_black'
    GRAVEYARD = 'terror_graveyard'
    SINGULARITY = 'energy_singularity'
    FF_PORTAL = 'ff_portal'
    CAVERN = 'shadow_cavern'
    MOLTENFURY = 'mount_moltenfury'
    FIRE_TEMPLE = 'fire_temple'
    FLAME_BRAZIER = 'flame_brazier'
    FIRE_UNIVERSE = 'fire_universe'
    SOUL_PORTAL = 'soul_portal'
    SOUL_TEMPLE = 'soul_temple'
    CRYPT = 'soul_crypt'
    HOLLOW = 'soul_zone'
    FORGE = 'soul_forge'
    FABRIC = 'fabric_universe'
    QUARK_PORTAL = 'quark_portal'
    QUARK_NEXUS = 'quark_nexus'
    QUANTUM_AETHER = 'quantum_aether'
    ASTRAL_OASIS = 'astral_oasis'
    PLANCK_SCOPE = 'planck_scope'
    TAPESTRY = 'dimensional_tapestry'
    HALLOWEEN = 'cursed_halloween'
    FARM_FIELD = 'farm_field'
    BUTTERFLY_FIELD = 'butterfly_field'
    VIAL = 'vial_of_life'
    DOOMED_FOREST = 'the_doomed_forest'
    
class Shortcut(str, Enum):
    DRAW_HUD = 'draw_hud'
    AFK_MODE = 'afk_mode'
    RESOURCE_FILTER = 'resource_draw_selection'
    HOUSE_CONSTRUCTOR = 'house_constructor'
    TRANSCEND_SELECTED = 'transcend_card'
    REFRESH_TRADES = 'refresh_trades'
    DAILY_REWARDS = 'daily_rewards'
    SCROLLS = 'scrolls'
    COMMUNITY_LEAVES = 'community_leaves'
    LEAFSCENSION = 'leafscension'
    LORE = 'lore'
    CHALLENGES = 'challenges'
    UNIQUE_LEAVES = 'unique_leaves'
    RELICS = 'relics'
    TRADING = 'trading'
    ITEMS = 'items'
    CHESTS = 'chests'
    PETS = 'pets'
    PRINTERS = 'printers'
    SUPPORTER = 'supporter_shop'
    BOT_PANEL = 'bot_control_panel'
    ACHIEVEMENTS = 'achievements'
    ARTIFACTS = 'artifacts'
    CARDS = 'cards'
    SKINS = 'skins'
    AREAS = 'areas'
    STATS = 'stats'
    MATERIALS = 'materials'
    PRESTIGE = 'prestige'
    BLC = 'blc'
    MLC = 'mlc'
    CONVERTERS = 'converters'
    FARMING = 'farming'
    ALCHEMY = 'alchemy'
    BORBVENTURES = 'borbventures'
    CRAFTING = 'crafting'
    DICE = 'dice'
    QUARKS = 'quarks'
    SHOP_TOOL = 'tool_shop'
    SHOP_BASIC = 'leaf_shop'
    SHOP_GOLD = 'gold_shop'
    SHOP_PLATINUM = 'platinum_shop'
    SHOP_BISMUTH = 'bismuth_shop'
    SHOP_COSMIC = 'cosmic_shop'
    VOID_SHOP = 'void_shop'
    SHOP_EXOTIC = 'exotic_shop'
    SHOP_CELEESTIAL = 'celestial_shop'
    SHOP_MYTHICAL = 'mythical_shop'
    SHOP_LAVA = 'lava_shop'
    SHOP_ICE = 'ice_shop'
    SHOP_OBSIDIAN = 'obsidian_shop'
    SHOP_SILICON = 'silicon_shop'
    SHOP_BENITOITE = 'benitoite_shop'
    SHOP_MOONSTONE = 'moonstone_shop'
    SHOP_SAND = 'sand_shop'
    SHOP_ANCIENT = 'ancient_shop'
    SHOP_GEM_LEAVES = 'gem_leaves_shop'
    SHOP_SACRED = 'sacred_shop'
    SHOP_BIOTITE = 'biotite_shop'
    SHOP_MALACHITE = 'malachite_shop'
    SHOP_HEMATITE = 'hematite_shop'
    SHOP_PLASMA = 'plasma_shop'
    SHOP_COAL = 'coal_shop'
    RED_SCIENCE = 'red_science'
    GREEN_SCIENCE = 'green_science'
    BLUE_SCIENCE = 'blue_science'
    MAGENTA_SCIENCE = 'magenta_science'
    ORANGE_SCIENCE = 'orange_science'
    BLACK_SCIENCE = 'black_science'
    STRANGE_SCIENCE = 'strange_science'
    SHOP_SILVER_TOKEN = 'silver_token_shop'
    SHOP_GOLD_TOKEN = 'gold_token_shop'
    SHOP_GEMS = 'shop_gems'
    SHOP_PRESTIGE = 'coins_shop'
    SHOP_BLC = 'blc_shop'
    SHOP_MLC = 'mlc_shop'
    SHOP_MULCH = 'mulch_shop'
    SHOP_ELECTRICAL = 'energy_electrical_shop'
    SHOP_SOUL = 'soul_shop'
    SHOP_EVENT = 'event_shop'
    BLAZING_SKULL = 'artifact_blazing_skull'
    ENCHANTED_FRUIT = 'artifact_fruit'
    SUITCASE = 'artifact_gold_suitcase'
    GRAV = 'artifact_gravity_ball'
    WIND = 'artifact_wind'
    WINGS = 'artifact_wings'
    VORTEX = 'artifact_vortex'
    VIOLIN = 'artifact_vital_violin'
    LANTERN = 'artifact_lantern'
    NATURE_ROD = 'artifact_nature_rod'
    LE = 'artifact_leafscension_exploit'
    ORB = 'artifact_orb'
    SEED_BAG = 'artifact_seed_bag'    
    CRAFT_SET_0 = 'load_craft_set_0'
    CRAFT_SET_1 = 'load_craft_set_1'
    CRAFT_SET_2 = 'load_craft_set_2'
    CRAFT_SET_3 = 'load_craft_set_3'
    CRAFT_SET_4 = 'load_craft_set_4'
    CRAFT_SET_5 = 'load_craft_set_5'
    CRAFT_SET_6 = 'load_craft_set_6'
    CRAFT_SET_7 = 'load_craft_set_7'
    LOADOUT_0 = 'load_loadout_0'
    LOADOUT_1 = 'load_loadout_1'
    LOADOUT_2 = 'load_loadout_2'
    LOADOUT_3 = 'load_loadout_3'
    LOADOUT_4 = 'load_loadout_4'
    LOADOUT_5 = 'load_loadout_5'
    LOADOUT_6 = 'load_loadout_6'
    LOADOUT_7 = 'load_loadout_7'
    
class Hotkey:
    HOTKEY_CODES = {
        8: "BACKSPACE",
        9: "TAB",
        160: "VK_LSHIFT",
        162: "VK_LCONTROL",
        164: "VK_LMENU",
        165: "VK_RMENU",
        92: "RWIN",
        163: "VK_RCONTROL",
        161: "VK_RSHIFT",
        12: "VK_NUMPAD5",
        13: "ENTER",
        16: "VK_SHIFT",
        17: "VK_CONTROL",
        18: "VK_MENU",
        19: "VK_PAUSE",
        20: "CAPSLOCK",
        32: "SPACE",
        33: "PGUP",
        34: "PGDN",
        35: "END",
        36: "HOME",
        37: "LEFT",
        38: "UP",
        39: "RIGHT",
        40: "DOWN",
        45: "INSERT",
        46: "DELETE",
        48: "0",
        49: "1",
        50: "2",
        51: "3",
        52: "4",
        53: "5",
        54: "6",
        55: "7",
        56: "8",
        57: "9",
        65: "a",
        66: "b",
        67: "c",
        68: "d",
        69: "e",
        70: "f",
        71: "g",
        72: "h",
        73: "i",
        74: "j",
        75: "k",
        76: "l",
        77: "m",
        78: "n",
        79: "o",
        80: "p",
        81: "q",
        82: "r",
        83: "s",
        84: "t",
        85: "u",
        86: "v",
        87: "w",
        88: "x",
        89: "y",
        90: "z",
        91: "LWIN",
        93: "VK_APPS",
        96: "VK_NUMPAD0",
        97: "VK_NUMPAD1",
        98: "VK_NUMPAD2",
        99: "VK_NUMPAD3",
        100: "VK_NUMPAD4",
        101: "VK_NUMPAD5",
        102: "VK_NUMPAD6",
        103: "VK_NUMPAD7",
        104: "VK_NUMPAD8",
        105: "VK_NUMPAD9",
        106: "VK_MULTIPLY",
        107: "VK_ADD",
        109: "VK_SUBTRACT",
        110: "VK_SEPARATOR",
        111: "VK_DIVIDE",
        112: "F1",
        113: "F2",
        114: "F3",
        115: "F4",
        116: "F5",
        117: "F6",
        118: "F7",
        119: "F8",
        120: "F9",
        121: "F10",
        122: "F11",
        123: "F12",
        144: "NUMLOCK",
        186: ";",
        187: "=",
        188: ":",
        189: "-",
        190: ".",
        191: "/",
        192: "'",
        219: "[",
        220: "^",
        221: "]",
        222: "'",
        223: "`",
        226: "\\"
    }
    
    def __init__(self, options: dict):
        self.settings = {}
        for k, v in options.items():
            if 'hotkey' == k.split('_')[0]:
                v = int(v['value']) if bool(v['value']) else -1
                self.settings.update({k: v})

    def use(self, shortcut: Shortcut, count: int = 1, pause: float = 0.05):
        settings_key = f'hotkey_{shortcut.value}'
        if self.settings[settings_key] == -1:
            raise KeyError('Missing key bind')
        key = self.HOTKEY_CODES[self.settings[settings_key]]
        for _ in range(count):
            send_keys(f'{{{key} down}} {{{key} up}}', pause) 

class LeafParser:
    SECRET = 'ke03m!5ng93nan7p24lyg343nml2o591'
    PATH = path.join(getenv('LOCALAPPDATA'), 'blow_the_leaves_away', 'save.dat')
    card_packs_drop_rarities = {
        Rarity.COMMON: {
            'drops': [Rarity.COMMON, Rarity.UNCOMMON, Rarity.RARE],
            'count': [1.425, 1.475, 0.1]
        },
        Rarity.RARE: {
            'drops': [Rarity.UNCOMMON, Rarity.RARE, Rarity.EPIC, Rarity.MYTHICAL],
            'count': [1.92375, 0.64125, 1.285, 0.15]
        },
        Rarity.LEGENDARY: {
            'drops': [Rarity.RARE, Rarity.EPIC, Rarity.MYTHICAL, Rarity.LEGENDARY],
            'count': [2.565, 0.855, 1.38, 0.2]
        }
    }
    ARTIFACT_CAP = 200

    def __init__(self) -> None:
        self.save = None
        self.options = None
        self.last_loaded = {'new': 0, 'old': 0}
        self.artifacts = {}
        self.cards = {
            rarity: {
                'regular': {},
                'boss': {}
            }
            for rarity in Rarity
        }
        for rarity in Rarity:
            for card in Regular:
                self.cards[rarity]['regular'].update({card: {'count': 0, 'transcend_cost': 0, 'locked': False, 'transcendable': False}})
            for card in Boss:
                self.cards[rarity]['boss'].update({card: {'count': 0, 'transcend_cost': 0, 'locked': False, 'transcendable': False}})
        self.cards_max = {}
        self.card_packs = {Rarity.COMMON: 0, Rarity.RARE: 0, Rarity.LEGENDARY: 0}
        self.trades = []
        self.resources = {}
        self.auto_teleport_area = None
        self.ss_kill_counter = 0
        self.gf_kill_counter = 0
        self._load_options()
    
    def parse_save(self) -> None:
        """Loads the current save and extract game data into dicts"""
        _time = self._load_save()
        if not bool(self.last_loaded['new']):
            raise ValueError('Empty save')
        else:
            print(f'Save loaded at: {_time}')
                
        _artifacts = self.save['profiles']['def']['artifacts']
        _cards = self.save['profiles']['def']['cards']
        _card_packs = self.save['profiles']['def']['cards_packs']
        _cards_transcensions = self.save['profiles']['def']['objects']['o_game']['data']['stats']['cards_transcensions']
        _areas = self.save['profiles']['def']['areas']
        self.trades = self.save['profiles']['def']['trades']
        self.trades.sort(reverse=True, key=lambda a: a['id'])
        self.resources = self.save['profiles']['def']['resources']
        self.gf_kill_counter = self.save['profiles']['def']['enemies']['boss_green_flame']['defeat_counter']
        self.ss_kill_counter = self.save['profiles']['def']['enemies']['boss_spectralseeker']['defeat_counter']
        
        for artifact in _artifacts:
            self.artifacts.update({Artifact(artifact): int(_artifacts[artifact]['count'])})
        
        for card in _cards:
            temp = card.split('_')
            _type = 'boss' if 'boss' in temp else 'regular'
            name = Boss('_'.join(temp[1:])) if 'boss' in temp else Regular('_'.join(temp[1:]))
            count, cost, locked = _cards[card]['count'], int(_cards[card]['transcended']) + 10, bool(_cards[card]['locked'])
            transcend = not locked and count >= cost
            self.cards[Rarity(temp[0])][_type][name].update({
                'count': count,
                'transcend_cost': cost,
                'locked': locked,
                'transcendable': transcend
            })
        
        for rarity, amount in _cards_transcensions.items():
            self.cards_max.update({Rarity(rarity): amount // 30 + 10 + int(self.save['profiles']['def']['shops']['cards_transcension']['cards_max_count']['count'])})
        
        for rarity in _card_packs:
            self.card_packs.update({Rarity(rarity): int(_card_packs[rarity]['count'])})
    
        for area in _areas:
            if _areas[area]['auto_teleport']:
                self.auto_teleport_area = Area(area)
                break  
    
    def get_trades(self) -> list:
        return self.trades

    def get_resource(self, name: str) -> float:
        """
        Returns the amount of a given resource.
        
        :param name: The name of the resource
        :return: Returns the resource called
        """
        return self.resources[name]['count']
    
    def get_artifact_count(self, artifact: Artifact) -> int:
        """
        Returns the amount of artifacts of the specified type
        
        :param artifact: The artifact enum, possible options in Artifact
        :return: Returns the amount of artifacts
        """
        if artifact not in Artifact:
            raise ValueError('Not an artifact')
       
        return self.artifacts.get(artifact)
    
    def get_card_count(self, rarity: Rarity, name: Regular | Boss) -> int:
        """
        Returns the amount of cards of the specified type
        
        :param rarity: The rarity of the card
        :param name: The name of the card, possible options in card_names, includes boss cards
        :return: Returns the amount of cards
        """
        if rarity not in Rarity and name not in Regular and name not in Boss:
            raise ValueError('Invalid card rarity or card name')
        
        _type = 'regular' if 'boss' != name[:4] else 'boss'
        return self.cards[rarity][_type][name].get('count')
    
    def get_card_locked(self, rarity: Rarity, name: Regular | Boss) -> bool:
        """
        Returns True if the specified card is locked, False otherwise
        
        :param rarity: The rarity of the card
        :param name: The name of the card, possible options in card_names, includes boss cards
        :return: Returns True if the specified card is locked, else False
        """
        if rarity not in Rarity and name not in Regular and name not in Boss:
            raise ValueError('Invalid card rarity or card name')
        
        _type = 'regular' if 'boss' != name[:4] else 'boss'
        return self.cards[rarity][_type][name].get('locked')
    
    def get_card_transcend_cost(self, rarity: Rarity, name: Regular | Boss) -> int:
        """
        Returns the amount of cards of a given card needed for transcension
        
        :param rarity: The rarity of the card
        :param name: The name of the card, possible options in card_names, includes boss cards
        :return: Returns the amount of cards needed for transcension
        """
        if rarity not in Rarity and name not in Regular and name not in Boss:
            raise ValueError('Invalid card rarity or card name')
        
        _type = 'regular' if 'boss' != name[:4] else 'boss'
        return self.cards[rarity][_type][name].get('transcend_cost')
    
    def get_rarity_transcend_count(self, rarity: Rarity) -> int:
        """
        Returns the number of cards in a given rarity that can be transcended
        
        :param rarity: The rarity of the cards
        :param boss: Enables boss transcension count (defaults to False)
        :return: Returns the numbert of cards that can be transcended
        """
        if rarity not in Rarity:
            raise ValueError

        return len([*filter(lambda x : x['transcendable'], self.cards[rarity]['regular'].values())])
    
    def get_pack_count(self, rarity: Rarity) -> int:
        if rarity not in self.card_packs:
            raise ValueError('Not a pack rarity')
        
        return self.card_packs.get(rarity)
    
    def get_current_player_area(self) -> str:
        if area := self.save['profiles']['def']['objects']['o_game']['data'].get('current_area_key'):
            return Area(area)
        return None
    
    def get_current_auto_teleport_area(self) -> str:
        return self.auto_teleport_area
    
    def _load_options(self) -> None:
        with open(path.join(getenv('localappdata'), 'blow_the_leaves_away', 'options.dat'), 'r') as file:
            self.options = json.load(file)
        
    def _load_save(self) -> str:
        with open(self.PATH, 'r') as encoded:
            decoded = str(b64decode(encoded.read()))[2:-43]
        
        self.save = json.loads(decoded)
        del decoded
        time_ = time.localtime()
        self.last_loaded['old'] = self.last_loaded['new']
        self.last_loaded['new'] = time.mktime(time_)
        return time.strftime("%d/%m/%Y %H:%M:%S", time_)
    
    def get_load_time(self, new=True):
        """Returns the latest save loaded time by default"""
        return self.last_loaded['new' if new else 'old']
    
    def __str__(self) -> str:
        return f'Save loaded and parsed at {time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(self.last_loaded["new"]))}'

class MacroController:
    game_path = 'D:\SteamLibrary\steamapps\common\Leaf Blower Revolution\game.exe'
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    locs = {
        'center': (960, 540),
        'card_tabs': {
            'packs': (310, 910),
            'transcend': (485, 910)
        },
        'areas': {
            Area.MOON: (1320, 805),
            Area.PLASMA_FOREST: (1320, 375),
            Area.CAVERN: (1320, 280),
            Area.FLAME_BRAZIER: (1320, 625),
            Area.FIRE_UNIVERSE: (1320, 765),
            Area.HOLLOW: (1320, 510)
        },
        'area_tabs': {
            'leaf_galaxy': (510, 910),
            'energy_belt': (950, 910),
            'fire_fields': (1150, 910),
            'soul_realm': (1350, 910),
            'favorite': (320, 910)
        },
        'menu_tabs': {
            'general': (310, 910)
        },
        'buttons': {
            'save': (670, 770),
            'transcend': (390, 310),
            'buy_legendary_packs': (1260, 735),
            'open_legendary_packs': (1320, 640),
            'buy_rare_packs': (840, 735),
            'open_rare_packs': (895, 640),
            'buy_common_packs': (420, 735),
            'open_common_packs': (475, 640),
            'boost_all': (),
            'collect_trades': (),
            'auto_refresh': (),
            'borbiana_jones': (1320, 280),
            'reset_spectralseeker': (780, 400)
        }
    }
    combination_hotkeys = {
        25000: ('VK_CONTROL', 'VK_SHIFT', 'VK_MENU'),
        2500: ('VK_CONTROL', 'VK_MENU'),
        1000: ('VK_SHIFT', 'VK_MENU')
        # 250: ('VK_CONTROL', 'VK_SHIFT'),
        # 100: ('VK_MENU',),
        # 25: ('VK_CONTROL',),
        # 10: ('VK_SHIFT',),
    }
    opening_times = {
        Rarity.COMMON: {
            1000: 0.5
        },
        Rarity.RARE: {
            1000: 0
        },
        Rarity.LEGENDARY: {
            1000: 0.75
        }
    }

    def __init__(self):
        self._app = Application(backend='uia').connect(path=MacroController.game_path)
        self._win = self._app.window(class_name='YYGameMakerYY')
        self.scheduler = sched.scheduler(time.monotonic, time.sleep)
        self.parser = LeafParser()
        self.detector = cv.SimpleBlobDetector_create(self._generate_detector_params())
        self.start_time = time.time()
        self.ff_tracker = {}
        self.general_tab_open = False
        self.current_area_tab = None
        self._running = False
        self.ss_farming = False
        
        self._win.maximize()
        self._win.wait('visible')
        self._win.set_focus()
        self.save_game()
        self.hotkeys = Hotkey(self.parser.options)
        self.ff_tracker['gf'] = self.parser.gf_kill_counter
        self.ff_tracker['ss'] = self.parser.ss_kill_counter

    def start(self):
        """ Schedule subroutines and runs them """
        self._running = True
        self.periodic(300, self.parser.parse_save, 2)
        self.periodic(5, self._task)
        self.scheduler.run()
    
    def stop(self):
        """ Stop all subroutines and exit application """
        self._running = False
        list(map(self.scheduler.cancel, self.scheduler.queue))
        exit(0)
        
    def periodic(self, interval, action, priority=1, actionargs=()):
        if self._running:
            self.scheduler.enter(interval, priority, self.periodic,
                                (interval, action, priority, actionargs))
            action(*actionargs)
        
    def _task(self):
        """ Contains the subroutines """
        le = bool(self.parser.get_artifact_count(Artifact.LE))
        lantern = bool(self.parser.get_artifact_count(Artifact.LANTERN))
        area = self.parser.get_current_player_area()
        dest = self.parser.get_current_auto_teleport_area()
        
        if not le:
            self.goto(Area.PLASMA_FOREST)
            self._collect_artifact(Artifact.LE)
        elif not lantern:
            self.goto(Area.MOON)
            self._collect_artifact(Artifact.LANTERN)
        elif self.ss_farming:
            start = time.monotonic()
            while time.monotonic() - start < 60:
                if self.ff_tracker['gf']:
                    self.kill_ss()
                else:
                    self.kill_gf()
        elif dest != area:
            self.goto(dest)
            self.save_game()
        else:
            match area:
                case Area.HOLLOW | Area.SINGULARITY:
                    start = time.monotonic()
                    while time.monotonic() - start < 30:
                        self.hotkeys.use(Shortcut.VIOLIN, 20, 0.005)
                        self.hotkeys.use(Shortcut.WIND)
                        self.hotkeys.use(Shortcut.GRAV)
                        self.hotkeys.use(Shortcut.VIOLIN, 20, 0.006)
                        time.sleep(1)
                case Area.CRYPT:
                    start = time.monotonic()
                    while time.monotonic() - start < 15:
                        self.hotkeys.use(Shortcut.WIND)
                        time.sleep(2)
                case Area.GLADE | Area.FARM_FIELD | Area.VIAL:
                    # pass
                    # if self.parser.get_pack_count(Rarity.COMMON):
                    #     self.open_packs(Rarity.COMMON)
                    # elif self.parser.get_pack_count(Rarity.RARE):
                    #     self.open_packs(Rarity.RARE)
                    if self.parser.get_pack_count(Rarity.LEGENDARY):
                      self.open_packs(Rarity.LEGENDARY)
                case Area.DICEY_MEADOWS:
                    pass
            
    def open_packs(self, rarity: Rarity, amount=0):
        """ 
        Open card packs of common, rare, or legendary rarity
        
        :param rarity: The pack rarity
        :param amount: The amount to be opened, defaults to all
        """    
        if rarity not in self.parser.card_packs_drop_rarities.keys():
            raise ValueError('Invalid rarity')
        
        self.hotkeys.use(Shortcut.CARDS)
        self._click(self.locs['card_tabs']['packs'])
        while (transcendable := self._can_transcend(rarity)) or self.parser.get_pack_count(rarity):
            if not (bool(self.parser.get_artifact_count((Artifact.LE))) and bool(self.parser.get_artifact_count(Artifact.LANTERN))):
                send_keys('{x down} {x up}')
                return
            elif transcendable:
                self.transcend(rarity)
            else:
                self._generate_open_seq(rarity)
                self.transcend(rarity)
                
        self.hotkeys.use(Shortcut.CARDS)
    
    def trade(self):
        queue = []
        
        pass
    
    def venture(self):
        pass
    
    # FIXME: fix gf kill calc
    def kill_gf(self):
        self.goto(Area.FLAME_BRAZIER)
        start = time.monotonic()
        while time.monotonic() - start <= 18:
            self.hotkeys.use(Shortcut.WIND)
            self.hotkeys.use(Shortcut.GRAV)
            time.sleep(0.5)
            if not self._has_spawned(Boss.GREEN_FLAME):
                break
        
        if self._has_spawned(Boss.GREEN_FLAME):
            self.goto(Area.CAVERN)
            return
        
        self.ff_tracker['gf'] += 1
        while not self._has_spawned(Boss.SPECTRALSEEKER):
            self.hotkeys.use(Shortcut.VIOLIN, 20, 0.005)
    
    # FIXME: fix ss kill calc
    def kill_ss(self):
        self.goto(Area.FIRE_UNIVERSE)
        card_bonus = 1 + sum([self.parser.get_card_count(rarity, Boss.SPECTRALSEEKER) for rarity in Rarity]) * 0.025
        # Damage = floor((1 + gf_kills * ss_vuln) * (1 + card_damage))
        damage =  math.floor((1 + self.ff_tracker['gf'] * self.parser.save['profiles']['def']['shops']['coal']['spectralseeker_green_flame_multiplier']['count']) * card_bonus)
        max_ss = 0 if damage < 90 else 1 if damage == 90 else (damage - 90) // 45
        
        if self.ff_tracker['ss'] < max_ss:
            start = time.monotonic()
            while time.monotonic() - start <= 18:
                self.hotkeys.use(Shortcut.WIND)
                self.hotkeys.use(Shortcut.GRAV)
                time.sleep(0.5)
                if not self._has_spawned(Boss.SPECTRALSEEKER):
                    break
                
            if self._has_spawned(Boss.SPECTRALSEEKER):
                self.goto(Area.CAVERN)
                return
            
            self.ff_tracker['gf'] = 0
            self.ff_tracker['ss'] += 1
        elif False:
            pass
        else:
            self.goto(Area.CAVERN)
            self._click(self.locs['buttons']['borbiana_jones'])
            self._click(self.locs['buttons']['reset_spectralseeker'])
            
            self.ff_tracker['gf'] = 0
            self.ff_tracker['ss'] = 0
            
        while not self._has_spawned(Boss.GREEN_FLAME):
            self.hotkeys.use(Shortcut.VIOLIN, 12, 0.005)
              
    def goto(self, dest: Area):
        """
        Go to the speicifed destination
        
        :param dest: The destination to travel to
        """
        self.hotkeys.use(Shortcut.AREAS)
        match dest:
            case self.parser.auto_teleport_area:
                if self.current_area_tab != 'favorite':
                    self._click(self.locs['area_tabs']['favorite'])
                    self.current_area_tab = 'favorite'
                send_keys('{SPACE down} {SPACE up}')
            case Area.MOON:
                if self.current_area_tab != 'leaf_galaxy':
                    self._click(self.locs['area_tabs']['leaf_galaxy'])
                    self.current_area_tab = 'leaf_galaxy'
                self._scroll(self.locs['areas'][Area.MOON], 14)
                self._click(self.locs['areas'][Area.MOON])
            case Area.PLASMA_FOREST:
                if self.current_area_tab != 'energy_belt':
                    self._click(self.locs['area_tabs']['energy_belt'])
                    self.current_area_tab = 'energy_belt'
                self._click(self.locs['areas'][Area.PLASMA_FOREST])
            case Area.CAVERN:
                if self.current_area_tab != 'fire_fields':
                    self._click(self.locs['area_tabs']['fire_fields'])
                    self.current_area_tab = 'fire_fields'
                self._click(self.locs['areas'][Area.CAVERN])
            case Area.FLAME_BRAZIER:
                if self.current_area_tab != 'fire_fields':
                    self._click(self.locs['area_tabs']['fire_fields'])
                    self.current_area_tab = 'fire_fields'
                self._click(self.locs['areas'][Area.FLAME_BRAZIER])
            case Area.FIRE_UNIVERSE:
                if self.current_area_tab != 'fire_fields':
                    self._click(self.locs['area_tabs']['fire_fields'])
                    self.current_area_tab = 'fire_fields'
                self._click(self.locs['areas'][Area.FIRE_UNIVERSE])
            case _:
                raise ValueError('Not a destination')
        self.hotkeys.use(Shortcut.AREAS)
         
    def save_game(self):
        """ Manually save the game through the menu and parse the save """
        send_keys('{ESC down} {ESC up}')
        if not self.general_tab_open:
            self._click(self.locs['menu_tabs']['general'], delay=250) # Click on General
            self.general_tab_open = True
        self._click(self.locs['buttons']['save'], delay=250) # Click on Save
        send_keys('{ESC down} {ESC up}', pause=0.25)
        self.parser.parse_save()
              
    def _can_transcend(self, rarity: Rarity):
        """
        Returns True if cards are ready to be transcended, False otherwise
        
        :param rarity: The rarity of the pack being opened
        :return: Returns True if cards can be transcended
        """
        if rarity not in self.parser.card_packs_drop_rarities.keys():
            raise ValueError('Invalid rarity')
        
        count = 0
        for r in self.parser.card_packs_drop_rarities[rarity]['drops']:
            temp = self.parser.get_rarity_transcend_count(r)
            count += temp
            if temp >= 25 or count >= 51:
                return True
        return False
       
    def transcend(self, rarity: Rarity):
        """ 
        This function transcends all cards
        
        :param rarity: The rarity of the packs opened
        """
        self._click(self.locs['card_tabs']['transcend'], delay=250)
        self._click(self.locs['buttons']['transcend'])
        send_keys('{SPACE down} {SPACE up}')
        
        if self.parser.options['show_reward_dialogs']['value']:
            self._click(self.locs['card_tabs']['packs'], count=2, delay=5)
            self._click(self.locs['buttons'][f'buy_{rarity}_packs'])
            self._click(self.locs['buttons'][f'open_{rarity}_packs'])
            self.hotkeys.use(Shortcut.CARDS)
            time.sleep(0.5)
            self.parser.parse_save()
        else:
            self.hotkeys.use(Shortcut.CARDS)
            self.save_game()
            self.hotkeys.use(Shortcut.CARDS)
            self._click(self.locs['card_tabs']['packs'], count=2, delay=5)
                              
    def _collect_artifact(self, name: Artifact):
        """
        This function collects artifact.
        
        :param name: The artifact's name
        """
        if name not in Artifact:
            raise ValueError('Invalid artifact')
        
        # Screen boundary box
        x0, y0, x1, y1 = (0, 23, 1920, 1050)
        # Values are in HSV
        lower_blue = (81,50,50) # 81, 50, 50
        upper_blue = (117,255,255) # 177, 255, 255
        
        ui_hidden = False
        cond = self.parser.ARTIFACT_CAP if name == Artifact.LE else 1
        duration = 18 if name == Artifact.LANTERN and not self.parser.save['profiles']['def']['events']['halloween']['active'] else 5
        self._click(button='right')
        while self.parser.get_artifact_count(name) < cond:
            start = time.monotonic()
            while time.monotonic() - start < duration:
                if not ui_hidden:
                    ui_hidden = True
                    self.hotkeys.use(Shortcut.DRAW_HUD, pause=0.25)
                
                screen = ImageGrab.grab(bbox=(x0, y0, x1, y1))
                screen = np.array(screen)
                hsv = cv.cvtColor(screen, cv.COLOR_BGR2HSV)
                mask = cv.inRange(hsv, lower_blue, upper_blue)
 
                if keypoints := self.detector.detect(mask):
                    locx, locy = keypoints[0].pt
                    mouse.move((int(locx), int(locy) + y0))
                    start = time.monotonic()                        
                self.use_artifact(Artifact.FRUIT, delay=250)
            
            if ui_hidden:
                self.hotkeys.use(Shortcut.DRAW_HUD, pause=0.25)
                ui_hidden = False
                
            self.save_game()
        mouse.move((960, 540))
        time.sleep(0.5)
        self._click((960, 540), button='right', delay=500)    
    
    def _generate_open_seq(self, rarity: Rarity):
        """
        Generate the opening subroutines needed for 25 cards of a rarity or 50 cards of any rarity to be transcendable.
        
        :param rarity: The rarity of the pack opened
        """
        transcend_costs = []
        total_transcend = 0
        min_pack = 1_000_000
        for r, count in zip(self.parser.card_packs_drop_rarities[rarity]['drops'], self.parser.card_packs_drop_rarities[rarity]['count']):
            temp = []
            cards = self.parser.cards[r]['regular'].values()
            rarity_transcend = self.parser.get_rarity_transcend_count(r)
            total_transcend += rarity_transcend
            if rarity_transcend >= 25 or total_transcend >= 51:
                return
            for cost in map(lambda x : int((x['transcend_cost'] - x['count']) * 40 / count), filter(lambda x : not x['transcendable'] and not x['locked'], cards)):
                insort(temp, cost)
                insort(transcend_costs, cost)
            min_pack = min(temp[25-rarity_transcend-1], min_pack)
            
        min_pack = math.ceil(min(transcend_costs[51-total_transcend-1], self.parser.get_pack_count(rarity), min_pack) / 1000) * 1000
        for k, values in self.combination_hotkeys.items():            
            if (res := min_pack//k) > 0:
                _time = self.opening_times[rarity][1000] if not self.parser.options['show_reward_dialogs']['value'] else self.opening_times[rarity][1000] * 1.18
                send_keys(' '.join(f'{{{v} down}}' for v in values))
                for _ in range(res):   
                    self._click(self.locs['buttons'][f'open_{rarity.value}_packs'], delay=250)
                    if self.parser.options['show_reward_dialogs']['value']:
                        self.hotkeys.use(Shortcut.CARDS)
                    time.sleep(_time)
                min_pack -= res * k
                send_keys(' '.join(f'{{{v} up}}' for v in values))
            
    @staticmethod          
    def _has_spawned(name: Boss) -> bool:
        """
        Check if a boss has been spawned.
        
        :param boss: The name of the boss
        """
        bosses_bbox = {
            Boss.GREEN_FLAME:    (1759, 760, 1840, 785),
            Boss.SPECTRALSEEKER: (1759, 805, 1840, 830)
        }
        
        ss = ImageGrab.grab(bbox=bosses_bbox[name])
        im = np.array(ss)
        im = cv.cvtColor(im, cv.COLOR_RGB2GRAY)
        ret, im = cv.threshold(im, 127, 255, cv.THRESH_BINARY)
        kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
        im = cv.filter2D(im, -1, kernel=kernel)

        return pytesseract.image_to_string(im).strip().lower() == "spawned"
    
    @staticmethod
    def _click(coords=(), button='left', count=1, delay=10):
        """
        Click at a specified coordinate. Possible buttons are 'left', 'right', 'middle'.
        
        :param coords: A tuple represeting the coordinate on screen, defaults to the current cursor position
        :param button: The mouse button to use, defaults to left
        :param count: The number of clicks
        :param delay: The delay after each click in milliseconds, defaults to 10ms
        """
        if not coords:
            coords = win32api.GetCursorPos()
        for _ in range(count):
            mouse.press(button, coords)
            mouse.release(button, coords)
            time.sleep(delay/1000)
    
    @staticmethod
    def _scroll(coords=(), dist=1, dir=-1, delay=20):
        """
        Do mouse scrolls.
        
        :param coords: The coordinates of the cursor, defaults to the current position
        :parm dist: The number of scrolls to make, defaults to 1
        :param dir: The direction to scroll, defaults to -1 (-1: DOWN, 1: UP)
        :param delay: The delay between scrolls in milliseconds, defaults to 20ms
        """
        if not coords:
            coords = win32api.GetCursorPos()
        for _ in range(dist):
            mouse.scroll(coords, dir)
            time.sleep(delay/1000)
    
    @staticmethod
    def _generate_detector_params():
        """ Generate the parameters for SimpleBlobDetector. """
        params = cv.SimpleBlobDetector_Params()
        params.filterByColor = True
        params.blobColor = 255
        params.filterByCircularity = False
        params.filterByConvexity = False
        params.filterByInertia = False
        params.filterByArea = True
        params.minArea = 10
        params.maxArea = 500
        params.minThreshold = 0
        params.maxThreshold = 255
        params.minDistBetweenBlobs = 20
        params.minRepeatability = 2 
        return params
    
    def __str__(self) -> str:
        return (f'\nStart time: {time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(self.start_time))}\n'
                f'End time: {time.strftime("%d/%m/%Y %H:%M:%S", time.localtime())}\n'
                f'Time elapsed: {timedelta(seconds=int(time.time()-self.start_time))}')

def main():
    try:
        controller = MacroController()
        controller.start()  
    except KeyboardInterrupt:
        print(controller)
        controller.stop()
        
if __name__ == '__main__':
    main()
