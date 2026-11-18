# -*- coding: utf-8 -*-
"""FamiGen Studio template library (pure data).

A template is a ready-to-send brief, not geometry: picking one fills the chat
composer and sets the family category, and Claude builds the family from it.
The user can edit the brief (sizes, finishes, types) before sending. Sizes are
typical catalogue values in millimetres; manufacturer names are examples of
the product class only.
"""

# (group, title, category, brief)
_ROWS = (
    # Architecture
    ('Architecture', 'Single flush door', 'Door',
     'Single flush door 900 × 2100 mm in a 50 mm hardwood frame, 44 mm leaf, white laminate '
     'leaf, brushed steel lever handle. Types: 800, 900 and 1000 mm wide.'),
    ('Architecture', 'Double glazed door', 'Door',
     'Pair of aluminium storefront doors 1800 × 2100 mm overall, medium stile 100 mm, full '
     'clear glass, dark bronze frame, vertical pull handles 600 mm.'),
    ('Architecture', 'Casement window', 'Window',
     'Two-pane side-hung casement window 1200 W × 1500 H mm, 60 mm aluminium frame, '
     'clear double glazing, 150 mm sill projection, grey frame.'),
    ('Architecture', 'Round column with capital', 'Columns',
     'Round concrete column Ø400 mm, 3000 mm high, square 600 × 600 × 300 mm capital and '
     'matching 600 × 600 × 200 mm base.'),
    # Furniture & casework
    ('Furniture & casework', 'Dining table', 'Furniture',
     'Dining table 2400 × 1200 × 750 mm, 40 mm oak top, four round black steel legs Ø60 mm '
     'set 100 mm in. Types: Oak and Walnut tops.'),
    ('Furniture & casework', 'Office desk', 'Furniture',
     'Office desk 1600 × 800 × 740 mm, 25 mm white melamine top, two black steel '
     'sled legs 60 × 20 mm, modesty panel 1400 × 350 mm.'),
    ('Furniture & casework', 'Lounge chair', 'Furniture',
     'Lounge chair 760 W × 820 D × 780 H mm, seat height 420 mm, grey fabric seat and back '
     'cushions, walnut frame.'),
    ('Furniture & casework', 'Base cabinet', 'Casework',
     'Kitchen base cabinet 600 W × 580 D × 870 H mm, 18 mm white carcass, one door, '
     '100 mm plinth, brushed steel bar handle. Types: 400, 600, 800 mm wide.'),
    ('Furniture & casework', 'Wall cabinet', 'Casework',
     'Wall cabinet 800 W × 350 D × 720 H mm, two doors, oak veneer, recessed handles.'),
    ('Furniture & casework', 'Reception counter', 'Casework',
     'Reception counter 2400 × 700 × 1100 mm with a 750 mm high writing shelf on the inside, '
     'white solid surface top, oak front panels.'),
    # Plumbing
    ('Plumbing', 'Wall-hung washbasin', 'Plumbing Fixture',
     'Wall-hung washbasin 600 × 450 × 160 mm, white ceramic, single tap hole, chrome mixer '
     'tap 160 mm high.'),
    ('Plumbing', 'Wall-hung WC', 'Plumbing Fixture',
     'Wall-hung WC pan 540 D × 360 W × 400 H mm, seat height 420 mm, white ceramic, '
     'flush plate 250 × 150 mm on the wall above.'),
    ('Plumbing', 'Shower tray with screen', 'Plumbing Fixture',
     'Shower tray 900 × 900 × 40 mm, white, with a 2000 mm high clear glass screen on two '
     'sides and a chrome rain head Ø250 mm.'),
    ('Plumbing', 'Electric water heater', 'Mechanical Equipment',
     'Vertical electric storage water heater Ø560 mm, 1250 mm high, white casing, top inlet '
     'and outlet stubs Ø22 mm, control box on the front.'),
    # MEP: HVAC
    ('MEP: HVAC', 'Rooftop air handling unit', 'Mechanical Equipment',
     'Rooftop AHU 3200 L × 1400 W × 1600 H mm, galvanised casing, two Ø630 mm fan outlets on '
     'top, 100 mm base frame, access doors on one side.'),
    ('MEP: HVAC', 'Fan coil unit (ceiling)', 'Mechanical Equipment',
     'Ceiling concealed fan coil unit 1000 × 600 × 250 mm, galvanised casing, supply spigot '
     '800 × 150 mm, return plenum on the back.'),
    ('MEP: HVAC', 'VAV box with reheat coil', 'Duct Accessory',
     'VAV box 600 L × 400 W × 300 H mm, Ø200 mm round inlet, 400 × 300 mm rectangular outlet, '
     'hot water reheat coil section 150 mm, control box on the side.'),
    ('MEP: HVAC', 'Fire damper', 'Duct Accessory',
     'Rectangular fire damper 600 × 400 mm, 300 mm long galvanised casing, 40 mm flanges, '
     'fusible link housing on top.'),
    ('MEP: HVAC', 'Split AC outdoor unit', 'Mechanical Equipment',
     'Split AC outdoor unit 900 W × 350 D × 700 H mm, white casing, Ø500 mm fan grille on the '
     'front, mounting feet 50 mm.'),
    # MEP: Fire protection
    ('MEP: Fire protection', 'Pendent sprinkler head', 'Sprinkler',
     'Pendent fire sprinkler head, 15 mm thread, 57 mm high, Ø55 mm deflector, red frangible '
     'glass bulb, chrome body, Ø80 mm escutcheon.'),
    ('MEP: Fire protection', 'Fire extinguisher cabinet', 'Specialty Equipment',
     'Recessed fire extinguisher cabinet 610 W × 250 D × 760 H mm, white steel tub, full glass '
     'door, 6 kg red extinguisher inside.'),
    ('MEP: Fire protection', 'Fire hose reel', 'Specialty Equipment',
     'Fire hose reel cabinet 800 × 300 × 800 mm, red steel, Ø600 mm reel, 30 m hose, '
     'glazed door.'),
    # MEP: Piping & supports
    ('MEP: Piping & supports', 'Flanged gate valve', 'Pipe Accessory',
     'Flanged gate valve DN100, 229 mm face to face, Ø220 mm flanges 24 mm thick, blue epoxy '
     'body, rising stem with Ø250 mm handwheel 420 mm above the pipe axis.'),
    ('MEP: Piping & supports', 'Quarter-turn ball valve', 'Pipe Accessory',
     'Full-port ball valve DN25, 90 mm long, brass body Ø45 mm, red lever handle 120 mm.'),
    ('MEP: Piping & supports', 'Butterfly valve', 'Pipe Accessory',
     'Wafer butterfly valve DN150, 56 mm face to face, Ø285 mm body, red gear operator box '
     '150 × 100 × 100 mm with a Ø200 mm handwheel.'),
    ('MEP: Piping & supports', 'Y-strainer', 'Pipe Accessory',
     'Flanged Y-strainer DN80, 310 mm face to face, Ø200 mm flanges, screen leg at 45° '
     'downwards 180 mm long, blue body.'),
    ('MEP: Piping & supports', 'End-suction pump', 'Mechanical Equipment',
     'Base-mounted end-suction pump set 1100 L × 450 W × 600 H mm: steel baseplate 60 mm, '
     'blue pump casing with DN80 suction and DN65 discharge, motor Ø280 mm.'),
    ('MEP: Piping & supports', 'Strut trapeze hanger', 'Generic Model',
     'Trapeze pipe hanger: 41 × 41 mm galvanised strut 900 mm long, hung on two M10 threaded '
     'rods 600 mm long.'),
    # MEP: Electrical
    ('MEP: Electrical', 'Distribution board', 'Electrical Equipment',
     'Surface-mounted distribution board 600 W × 200 D × 900 H mm, grey steel enclosure, '
     'hinged door with a handle, cable gland plate on top.'),
    ('MEP: Electrical', 'Dry-type transformer', 'Electrical Equipment',
     'Ventilated dry-type transformer 1500 W × 900 D × 1800 H mm, grey enclosure, louvres on '
     'the front and sides, lifting eyes on top.'),
    ('MEP: Electrical', 'Safety switch', 'Electrical Equipment',
     'Disconnect safety switch 250 W × 150 D × 400 H mm, grey enclosure, side operating '
     'handle 120 mm.'),
    ('MEP: Electrical', 'Wall socket outlet', 'Electrical Fixture',
     'Double socket outlet 146 × 86 × 40 mm, white plastic face plate, two switched sockets.'),
    ('MEP: Electrical', 'Junction box', 'Electrical Fixture',
     'Surface junction box 150 × 150 × 100 mm, grey, four knockouts Ø25 mm.'),
    # Lighting
    ('Lighting', 'Pendant light', 'Lighting Fixture',
     'Pendant light: Ø450 mm aluminium dome shade, matt black, 1500 mm cable, Ø120 mm ceiling '
     'canopy.'),
    ('Lighting', 'Recessed downlight', 'Lighting Fixture',
     'Recessed LED downlight Ø200 mm trim, 110 mm body depth, white trim, opal diffuser.'),
    ('Lighting', 'Linear LED fixture', 'Lighting Fixture',
     'Suspended linear LED luminaire 1500 × 70 × 80 mm, anodised aluminium, opal diffuser, '
     'two suspension cables 1000 mm.'),
    # Site & entourage
    ('Site & entourage', 'Bike rack', 'Site',
     'Outdoor bike rack, galvanised steel, five inverted-U hoops Ø48 mm, 800 mm high, 600 mm '
     'apart, on a 3000 mm base rail.'),
    ('Site & entourage', 'Concrete bollard', 'Site',
     'Concrete bollard Ø200 mm, 900 mm high, 20 mm chamfer at the top, grey concrete.'),
    ('Site & entourage', 'Park bench', 'Site',
     'Park bench 1800 × 650 × 800 mm, seat height 450 mm, timber slats, black cast iron ends.'),
    ('Site & entourage', 'Planter box', 'Entourage',
     'Square planter 800 × 800 × 600 mm, 30 mm corten steel, with a Ø1200 mm shrub crown '
     'above.'),
)


def templates():
    """[{id, group, title, category, prompt}] in display order."""
    out = []
    for index, (group, title, category, brief) in enumerate(_ROWS):
        out.append({'id': 't{:02d}'.format(index + 1), 'group': group, 'title': title,
                    'category': category, 'prompt': brief})
    return out


def groups():
    """Group names in display order."""
    seen = []
    for row in _ROWS:
        if row[0] not in seen:
            seen.append(row[0])
    return seen
