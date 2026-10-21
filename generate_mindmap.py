import os
import glob
import json
import openpyxl
from difflib import SequenceMatcher
from collections import defaultdict


# ============================================================
# CONFIGURATION
# ============================================================

SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = "index.html"

START_ROW = 20
END_ROW = 100

SIMILARITY_THRESHOLD = 0.8

MODULE_COLUMNS = 3


# ============================================================
# HELPERS
# ============================================================

def clean(value):
    if value is None:
        return ""
    return str(value).strip()


def similarity(a, b):
    if not a or not b:
        return 0

    return SequenceMatcher(
        None,
        str(a),
        str(b)
    ).ratio()


# ============================================================
# READ EXCEL FILES
# ============================================================

os.chdir(SOURCE_DIR)

excel_files = glob.glob("*.xlsx")

print()
print("=" * 70)
print("COLLABORATION MIND MAP GENERATOR")
print("=" * 70)
print()

all_modules = []


for filename in excel_files:

    print("Reading:", filename)

    try:

        workbook = openpyxl.load_workbook(
            filename,
            data_only=True
        )

        sheet = workbook.active

        category = (
            clean(sheet["D10"].value)
            or "Category"
        )

        subcategory = (
            clean(sheet["B10"].value)
            or "Sub-Category"
        )

        stage = (
            clean(sheet["F10"].value)
            or "Section"
        )

        path = (
            f"{category} > "
            f"{subcategory} > "
            f"{stage}"
        )

        for row in sheet.iter_rows(
            min_row=START_ROW,
            max_row=END_ROW,
            min_col=3,
            max_col=8
        ):

            subtopic = clean(
                row[3].value
            )

            if not subtopic:
                continue

            status = clean(
                row[0].value
            ).lower()

            description = clean(
                row[5].value
            )

            all_modules.append({

                "category": category,

                "subcategory": subcategory,

                "stage": stage,

                "path": path,

                "status": status,

                "subtopic": subtopic,

                "description": description,

                "related_to": "",

                "related_path": ""

            })

    except Exception as error:

        print(
            "ERROR reading",
            filename,
            ":",
            error
        )


print()
print(
    "Content items loaded:",
    len(all_modules)
)


# ============================================================
# FIND RELATED CONTENT
# ============================================================

for index, item in enumerate(all_modules):

    for other_index, other in enumerate(all_modules):

        if index == other_index:
            continue

        if not item["description"]:
            continue

        if not other["description"]:
            continue

        score = similarity(
            item["description"],
            other["description"]
        )

        if score >= SIMILARITY_THRESHOLD:

            item["related_to"] = (
                other["subtopic"]
            )

            item["related_path"] = (
                other["path"]
            )

            break


# ============================================================
# BUILD HIERARCHY
# ============================================================

hierarchy = defaultdict(
    lambda: defaultdict(
        lambda: defaultdict(list)
    )
)


for item in all_modules:

    hierarchy[
        item["category"]
    ][
        item["subcategory"]
    ][
        item["stage"]
    ].append(item)


# ============================================================
# NODE DATA
# ============================================================

nodes = []
links = []

next_node_id = 0


def add_node(
    label,
    level,
    parent=None,
    description="",
    status="",
    related_to="",
    related_path=""
):

    global next_node_id

    node = {

        "id": next_node_id,

        "label": str(label),

        "level": level,

        "parent": parent,

        "description": str(
            description or ""
        ),

        "status": str(
            status or ""
        ),

        "related_to": str(
            related_to or ""
        ),

        "related_path": str(
            related_path or ""
        )

    }

    nodes.append(node)

    if parent is not None:

        links.append({

            "source": parent,

            "target": next_node_id

        })

    created_id = next_node_id

    next_node_id += 1

    return created_id


# ============================================================
# ROOT
# ============================================================

root_id = add_node(
    "Collaboration Map",
    0
)


# ============================================================
# BUILD TREE
# ============================================================

for category, subcategories in hierarchy.items():

    category_id = add_node(
        category,
        1,
        root_id
    )

    for subcategory, stages in subcategories.items():

        subcategory_id = add_node(
            subcategory,
            2,
            category_id
        )

        for stage, modules in stages.items():

            stage_id = add_node(
                stage,
                3,
                subcategory_id
            )

            for module in modules:

                add_node(

                    module["subtopic"],

                    4,

                    stage_id,

                    module["description"],

                    module["status"],

                    module["related_to"],

                    module["related_path"]

                )


print(
    "Nodes created:",
    len(nodes)
)

print(
    "Connections created:",
    len(links)
)


# ============================================================
# JSON
# ============================================================

nodes_json = json.dumps(
    nodes,
    ensure_ascii=False
)

links_json = json.dumps(
    links,
    ensure_ascii=False
)


# ============================================================
# HTML
# ============================================================

HTML = r'''
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
>

<title>
Collaboration Mind Map
</title>


<style>

* {
    box-sizing: border-box;
}


html,
body {

    margin: 0;
    padding: 0;

    width: 100%;
    height: 100%;

    overflow: hidden;

    font-family:
        Arial,
        Helvetica,
        sans-serif;

    background:
        #f3f6f8;
}


/* ============================================================
   HEADER
   ============================================================ */

#header {

    position: fixed;

    top: 0;
    left: 0;
    right: 0;

    height: 62px;

    background:
        #005073;

    color: white;

    display: flex;

    align-items: center;

    justify-content: space-between;

    padding:
        0 18px;

    z-index: 10000;
}


#header h1 {

    margin: 0;

    font-size: 21px;
}


#buttons {

    display: flex;

    gap: 6px;
}


#buttons button {

    border: none;

    border-radius: 5px;

    background: white;

    color: #005073;

    padding:
        8px 14px;

    font-weight: bold;

    cursor: pointer;
}


#buttons button:hover {

    background:
        #e2f0f4;
}


/* ============================================================
   MAP
   ============================================================ */

#map {

    position: fixed;

    top: 62px;
    left: 0;
    right: 0;
    bottom: 0;

    overflow: hidden;

    cursor: grab;
}


#map.dragging {

    cursor: grabbing;
}


#world {

    position: absolute;

    left: 0;
    top: 0;

    width: 1px;
    height: 1px;

    transform-origin:
        0 0;
}


/* ============================================================
   SVG
   ============================================================ */

#line-layer {

    position: absolute;

    left: 0;
    top: 0;

    width: 1px;
    height: 1px;

    overflow: visible;

    pointer-events: none;
}


.line {

    fill: none;

    stroke-linecap: round;
}


.line1 {
    stroke: #005073;
    stroke-width: 7;
}


.line2 {
    stroke: #4d91a8;
    stroke-width: 5;
}


.line3 {
    stroke: #8eafb9;
    stroke-width: 4;
}


.line4 {
    stroke: #c4cdd1;
    stroke-width: 2.5;
}


/* ============================================================
   NODES
   ============================================================ */

.node {

    position: absolute;

    transform:
        translate(
            -50%,
            -50%
        );

    border-radius: 50%;

    display: flex;

    align-items: center;

    justify-content: center;

    text-align: center;

    padding: 10px;

    cursor: pointer;

    user-select: none;

    transition:
        box-shadow .15s ease,
        transform .15s ease;
}


.node:hover {

    transform:
        translate(
            -50%,
            -50%
        )
        scale(1.06);
}


.node.hidden {

    display: none;
}


/* ============================================================
   ROOT
   ============================================================ */

.node.level0 {

    width: 190px;
    height: 190px;

    background:
        #005073;

    color: white;

    border:
        8px solid white;

    font-size: 22px;

    font-weight: bold;

    z-index: 5000;

    box-shadow:
        0 8px 30px
        rgba(
            0,
            0,
            0,
            .30
        );
}


.node.level0.collapsed {

    box-shadow:
        0 0 0 12px
        rgba(
            0,
            80,
            115,
            .18
        ),
        0 8px 30px
        rgba(
            0,
            0,
            0,
            .30
        );
}


/* ============================================================
   CATEGORY
   ============================================================ */

.node.level1 {

    width: 155px;
    height: 155px;

    background:
        #0076a8;

    color: white;

    border:
        6px solid white;

    font-size: 15px;

    font-weight: bold;

    z-index: 4000;
}


/* ============================================================
   SUBCATEGORY
   ============================================================ */

.node.level2 {

    width: 130px;
    height: 130px;

    background:
        #5799af;

    color: white;

    border:
        5px solid white;

    font-size: 13px;

    font-weight: bold;

    z-index: 3500;
}


/* ============================================================
   STAGE
   ============================================================ */

.node.level3 {

    width: 140px;
    height: 140px;

    background:
        #dcebef;

    color:
        #003d52;

    border:
        6px solid white;

    font-size: 12px;

    font-weight: bold;

    font-style: italic;

    z-index: 4500;
}


/* ============================================================
   MODULE
   ============================================================ */

.node.level4 {

    width: 110px;
    height: 110px;

    background:
        white;

    color:
        #333;

    border:
        4px solid #c4cdd1;

    font-size: 10.5px;

    z-index: 3000;

    box-shadow:
        0 3px 10px
        rgba(
            0,
            0,
            0,
            .12
        );
}


.node.level4.mandatory {

    border-color:
        #d9534f;
}


.node.level4.optional {

    border-color:
        #5cb85c;
}


.node.level4.related {

    border-color:
        #6f42c1;

    background:
        #f8f5ff;
}


/* ============================================================
   TOOLTIP
   ============================================================ */

#tooltip {

    position: fixed;

    display: none;

    max-width: 420px;

    background:
        rgba(
            25,
            25,
            25,
            .96
        );

    color: white;

    padding: 15px;

    border-radius: 8px;

    z-index: 20000;

    pointer-events: none;

    font-size: 13px;

    line-height: 1.45;

    box-shadow:
        0 5px 25px
        rgba(
            0,
            0,
            0,
            .35
        );
}


.tooltip-title {

    font-size: 16px;

    font-weight: bold;

    margin-bottom: 8px;
}


.tooltip-related {

    border-top:
        1px solid #666;

    margin-top: 10px;

    padding-top: 10px;
}


/* ============================================================
   LEGEND
   ============================================================ */

#legend {

    position: fixed;

    left: 18px;
    bottom: 18px;

    background:
        white;

    padding:
        12px 15px;

    border-radius: 8px;

    z-index: 10000;

    font-size: 12px;

    box-shadow:
        0 3px 15px
        rgba(
            0,
            0,
            0,
            .15
        );
}


.legend-title {

    font-weight: bold;

    margin-bottom: 7px;
}


.legend-row {

    margin: 5px 0;
}


.dot {

    display: inline-block;

    width: 12px;
    height: 12px;

    border-radius: 50%;

    margin-right: 6px;
}


.red {
    background: #d9534f;
}


.green {
    background: #5cb85c;
}


.purple {
    background: #6f42c1;
}


/* ============================================================
   INSTRUCTIONS
   ============================================================ */

#instructions {

    position: fixed;

    right: 18px;
    bottom: 18px;

    background:
        white;

    padding:
        10px 14px;

    border-radius: 7px;

    z-index: 10000;

    font-size: 11px;

    color:
        #555;

    box-shadow:
        0 3px 12px
        rgba(
            0,
            0,
            0,
            .12
        );
}

</style>

</head>


<body>


<div id="header">

    <h1>
        Collaboration Mind Map
    </h1>

    <div id="buttons">

        <button id="reset">
            Reset
        </button>

        <button id="minus">
            −
        </button>

        <button id="plus">
            +
        </button>

        <button id="fit">
            Fit
        </button>

    </div>

</div>


<div id="map">

    <div id="world">

        <svg id="line-layer">

            <g id="lines"></g>

        </svg>

    </div>

</div>


<div id="tooltip"></div>


<div id="legend">

    <div class="legend-title">
        Content Status
    </div>

    <div class="legend-row">

        <span class="dot red"></span>
        Mandatory

    </div>

    <div class="legend-row">

        <span class="dot green"></span>
        Optional

    </div>

    <div class="legend-row">

        <span class="dot purple"></span>
        Related

    </div>

</div>


<div id="instructions">

    Click nodes to expand/collapse
    &nbsp; | &nbsp;
    Drag background
    &nbsp; | &nbsp;
    Scroll to zoom

</div>


<script>


// ============================================================
// DATA
// ============================================================

const nodes =
    __NODES__;


const links =
    __LINKS__;


// ============================================================
// DOM
// ============================================================

const map =
    document.getElementById(
        "map"
    );


const world =
    document.getElementById(
        "world"
    );


const lineGroup =
    document.getElementById(
        "lines"
    );


const tooltip =
    document.getElementById(
        "tooltip"
    );


// ============================================================
// STATE
// ============================================================

let rootOpen = false;

let zoom = 0.65;

let panX = 0;

let panY = 0;


// ============================================================
// LOOKUPS
// ============================================================

const byId = {};

const children = {};


nodes.forEach(
    function(node) {

        byId[node.id] =
            node;

        children[node.id] =
            [];

    }
);


links.forEach(
    function(link) {

        children[
            link.source
        ].push(
            link.target
        );

    }
);


// ============================================================
// OPEN STATE
// ============================================================

const open = {};


nodes.forEach(
    function(node) {

        open[node.id] =
            false;

    }
);


// ============================================================
// ELEMENTS
// ============================================================

const elements = {};


// ============================================================
// CREATE NODE ELEMENTS
// ============================================================

nodes.forEach(
    function(node) {

        const element =
            document.createElement(
                "div"
            );


        element.className =
            "node level" +
            node.level;


        element.textContent =
            node.label;


        element.dataset.id =
            node.id;


        // ----------------------------------------------------
        // MODULE STATUS
        // ----------------------------------------------------

        if (
            node.level === 4
        ) {

            const status =
                (
                    node.status ||
                    ""
                ).toLowerCase();


            if (
                status.includes(
                    "mandatory"
                )
            ) {

                element.classList.add(
                    "mandatory"
                );

            }


            if (
                status.includes(
                    "optional"
                )
            ) {

                element.classList.add(
                    "optional"
                );

            }


            if (
                node.related_to
            ) {

                element.classList.add(
                    "related"
                );

            }

        }


        world.appendChild(
            element
        );


        elements[node.id] =
            element;


        // ====================================================
        // CLICK
        // ====================================================

        element.addEventListener(
            "click",
            function(event) {

                event.preventDefault();

                event.stopPropagation();


                const id =
                    Number(
                        element.dataset.id
                    );


                const node =
                    byId[id];


                console.log(
                    "NODE CLICKED:",
                    id,
                    node.label,
                    "LEVEL:",
                    node.level
                );


                // --------------------------------------------
                // ROOT
                // --------------------------------------------

                if (
                    id === 0
                ) {

                    rootOpen =
                        !rootOpen;


                    console.log(
                        "ROOT:",
                        rootOpen
                    );


                    if (
                        !rootOpen
                    ) {

                        nodes.forEach(
                            function(n) {

                                open[n.id] =
                                    false;

                            }
                        );

                    }


                    render();

                    return;

                }


                // --------------------------------------------
                // LEAF / MODULE
                // --------------------------------------------

                if (
                    children[id].length === 0
                ) {

                    console.log(
                        "LEAF NODE - NO CHILDREN"
                    );

                    return;

                }


                // --------------------------------------------
                // TOGGLE THIS NODE
                // --------------------------------------------

                open[id] =
                    !open[id];


                console.log(
                    "NODE OPEN:",
                    node.label,
                    open[id]
                );


                // --------------------------------------------
                // IMPORTANT:
                //
                // When opening a node, don't automatically
                // close anything else. This allows multiple
                // stages to remain open simultaneously.
                // --------------------------------------------

                render();

            }
        );


        // ----------------------------------------------------
        // PREVENT DRAGGING FROM NODE
        // ----------------------------------------------------

        element.addEventListener(
            "pointerdown",
            function(event) {

                event.stopPropagation();

            }
        );


        // ----------------------------------------------------
        // TOOLTIP
        // ----------------------------------------------------

        element.addEventListener(
            "mouseenter",
            function(event) {

                if (
                    node.level === 4
                ) {

                    showTooltip(
                        node,
                        event
                    );

                }

            }
        );


        element.addEventListener(
            "mousemove",
            function(event) {

                if (
                    node.level === 4
                ) {

                    moveTooltip(
                        event
                    );

                }

            }
        );


        element.addEventListener(
            "mouseleave",
            function() {

                tooltip.style.display =
                    "none";

            }
        );

    }
);


// ============================================================
// CONNECTIONS
// ============================================================

const lineElements = {};


links.forEach(
    function(link) {

        const path =
            document.createElementNS(
                "http://www.w3.org/2000/svg",
                "path"
            );


        path.classList.add(
            "line"
        );


        path.classList.add(
            "line" +
            byId[
                link.target
            ].level
        );


        lineGroup.appendChild(
            path
        );


        lineElements[
            link.source +
            "-" +
            link.target
        ] =
            path;

    }
);


// ============================================================
// VISIBILITY
// ============================================================

function visible(
    nodeId
) {

    // --------------------------------------------------------
    // ROOT
    // --------------------------------------------------------

    if (
        nodeId === 0
    ) {

        return true;

    }


    // --------------------------------------------------------
    // ROOT CLOSED
    // --------------------------------------------------------

    if (
        !rootOpen
    ) {

        return false;

    }


    let node =
        byId[nodeId];


    // --------------------------------------------------------
    // WALK TOWARD ROOT
    // --------------------------------------------------------

    while (
        node.parent !== null
    ) {

        const parentId =
            node.parent;


        // --------------------------------------------
        // DIRECT CHILD OF ROOT
        // --------------------------------------------

        if (
            parentId === 0
        ) {

            return rootOpen;

        }


        // --------------------------------------------
        // ALL OTHER LEVELS
        // --------------------------------------------

        if (
            open[parentId] !== true
        ) {

            return false;

        }


        node =
            byId[parentId];

    }


    return true;

}


// ============================================================
// BRANCH SIZE
// ============================================================

function branchSize(
    nodeId
) {

    const kids =
        children[nodeId];


    if (
        kids.length === 0
    ) {

        return 1;

    }


    if (
        !open[nodeId]
    ) {

        return 1;

    }


    let total = 0;


    kids.forEach(
        function(id) {

            total +=
                branchSize(id);

        }
    );


    return Math.max(
        1,
        total
    );

}


// ============================================================
// LAYOUT
// ============================================================

function calculateLayout() {

    byId[0].x = 0;

    byId[0].y = 0;


    if (
        !rootOpen
    ) {

        return;

    }


    const categories =
        children[0];


    const right =
        categories.filter(
            function(id, index) {

                return index % 2 === 0;

            }
        );


    const left =
        categories.filter(
            function(id, index) {

                return index % 2 !== 0;

            }
        );


    layoutCategories(
        right,
        1
    );


    layoutCategories(
        left,
        -1
    );

}


// ============================================================
// CATEGORY LAYOUT
// ============================================================

function layoutCategories(
    ids,
    direction
) {

    if (
        ids.length === 0
    ) {

        return;

    }


    const spacing =
        400;


    let totalUnits = 0;


    ids.forEach(
        function(id) {

            totalUnits +=
                branchSize(id);

        }
    );


    const totalHeight =
        (
            totalUnits - 1
        ) *
        spacing;


    let y =
        -totalHeight / 2;


    ids.forEach(
        function(id) {

            const size =
                branchSize(id);


            const height =
                (
                    size - 1
                ) *
                spacing;


            const node =
                byId[id];


            node.x =
                direction *
                390;


            node.y =
                y +
                height / 2;


            y +=
                height +
                spacing;


            layoutChildren(
                id,
                direction
            );

        }
    );

}


// ============================================================
// CHILDREN LAYOUT
// ============================================================

function layoutChildren(
    parentId,
    direction
) {

    if (
        !open[parentId]
    ) {

        return;

    }


    const parent =
        byId[parentId];


    const kids =
        children[parentId];


    if (
        kids.length === 0
    ) {

        return;

    }


    // ========================================================
    // STAGE
    //
    // This is the important fix.
    //
    // A stage is LEVEL 3.
    // When it is open, its LEVEL 4 modules are explicitly
    // laid out in a 3-column grid.
    // ========================================================

    if (
        parent.level === 3
    ) {

        layoutModules(
            parent,
            kids,
            direction
        );


        return;

    }


    // ========================================================
    // CATEGORY / SUBCATEGORY
    // ========================================================

    const spacing =
        parent.level === 1
            ? 330
            : 280;


    let totalUnits = 0;


    kids.forEach(
        function(id) {

            totalUnits +=
                branchSize(id);

        }
    );


    const totalHeight =
        (
            totalUnits - 1
        ) *
        spacing;


    let y =
        parent.y -
        totalHeight / 2;


    kids.forEach(
        function(id) {

            const size =
                branchSize(id);


            const height =
                (
                    size - 1
                ) *
                spacing;


            const child =
                byId[id];


            child.x =
                parent.x +
                direction *
                (
                    parent.level === 1
                        ? 285
                        : 245
                );


            child.y =
                y +
                height / 2;


            y +=
                height +
                spacing;


            layoutChildren(
                id,
                direction
            );

        }
    );

}


// ============================================================
// MODULE GRID
// ============================================================

function layoutModules(
    stage,
    moduleIds,
    direction
) {

    const columns =
        3;


    const horizontalGap =
        140;


    const verticalGap =
        130;


    // Keep modules well clear of the stage.

    const startDistance =
        235;


    const rows =
        Math.ceil(
            moduleIds.length /
            columns
        );


    const totalHeight =
        (
            rows - 1
        ) *
        verticalGap;


    const startY =
        stage.y -
        totalHeight / 2;


    moduleIds.forEach(
        function(id, index) {

            const module =
                byId[id];


            const row =
                Math.floor(
                    index /
                    columns
                );


            const column =
                index %
                columns;


            module.x =
                stage.x +
                direction *
                (
                    startDistance +
                    (
                        column *
                        horizontalGap
                    )
                );


            module.y =
                startY +
                (
                    row *
                    verticalGap
                );

        }
    );

}


// ============================================================
// CONNECTION PATH
// ============================================================

function makePath(
    source,
    target
) {

    const direction =
        target.x >= source.x
            ? 1
            : -1;


    const distance =
        Math.abs(
            target.x -
            source.x
        );


    const curve =
        Math.max(
            70,
            distance * 0.4
        );


    const c1x =
        source.x +
        direction *
        curve;


    const c2x =
        target.x -
        direction *
        curve;


    return (
        "M " +
        source.x +
        " " +
        source.y +
        " C " +
        c1x +
        " " +
        source.y +
        ", " +
        c2x +
        " " +
        target.y +
        ", " +
        target.x +
        " " +
        target.y
    );

}


// ============================================================
// RENDER
// ============================================================

function render() {

    calculateLayout();


    // --------------------------------------------------------
    // NODES
    // --------------------------------------------------------

    nodes.forEach(
        function(node) {

            const element =
                elements[node.id];


            element.style.left =
                node.x + "px";


            element.style.top =
                node.y + "px";


            element.classList.toggle(
                "hidden",
                !visible(node.id)
            );

        }
    );


    // --------------------------------------------------------
    // CONNECTIONS
    // --------------------------------------------------------

    links.forEach(
        function(link) {

            const source =
                byId[
                    link.source
                ];


            const target =
                byId[
                    link.target
                ];


            const path =
                lineElements[
                    link.source +
                    "-" +
                    link.target
                ];


            const show =
                visible(source.id) &&
                visible(target.id);


            path.style.display =
                show
                    ? "block"
                    : "none";


            if (
                show
            ) {

                path.setAttribute(
                    "d",
                    makePath(
                        source,
                        target
                    )
                );

            }

        }
    );


    applyTransform();

}


// ============================================================
// TRANSFORM
// ============================================================

function applyTransform() {

    world.style.transform =
        "translate(" +
        panX +
        "px, " +
        panY +
        "px) scale(" +
        zoom +
        ")";

}


// ============================================================
// CENTRE
// ============================================================

function centre() {

    panX =
        map.clientWidth / 2;


    panY =
        map.clientHeight / 2;


    applyTransform();

}


// ============================================================
// RESET
// ============================================================

function resetMap() {

    rootOpen =
        false;


    nodes.forEach(
        function(node) {

            open[node.id] =
                false;

        }
    );


    elements[0]
        .classList.add(
            "collapsed"
        );


    zoom =
        0.65;


    centre();

    render();

}


// ============================================================
// FIT
// ============================================================

function fit() {

    zoom =
        0.45;


    centre();

    render();

}


// ============================================================
// BUTTONS
// ============================================================

document
    .getElementById("plus")
    .addEventListener(
        "click",
        function(event) {

            event.stopPropagation();


            zoom =
                Math.min(
                    3,
                    zoom * 1.15
                );


            applyTransform();

        }
    );


document
    .getElementById("minus")
    .addEventListener(
        "click",
        function(event) {

            event.stopPropagation();


            zoom =
                Math.max(
                    .2,
                    zoom / 1.15
                );


            applyTransform();

        }
    );


document
    .getElementById("fit")
    .addEventListener(
        "click",
        function(event) {

            event.stopPropagation();

            fit();

        }
    );


document
    .getElementById("reset")
    .addEventListener(
        "click",
        function(event) {

            event.stopPropagation();

            resetMap();

        }
    );


// ============================================================
// DRAGGING
// ============================================================

let dragging = false;

let startX = 0;

let startY = 0;

let initialPanX = 0;

let initialPanY = 0;


map.addEventListener(
    "pointerdown",
    function(event) {

        if (
            event.target.closest(
                ".node"
            )
        ) {

            return;

        }


        dragging =
            true;


        startX =
            event.clientX;


        startY =
            event.clientY;


        initialPanX =
            panX;


        initialPanY =
            panY;


        map.classList.add(
            "dragging"
        );


        map.setPointerCapture(
            event.pointerId
        );

    }
);


map.addEventListener(
    "pointermove",
    function(event) {

        if (
            !dragging
        ) {

            return;

        }


        panX =
            initialPanX +
            (
                event.clientX -
                startX
            );


        panY =
            initialPanY +
            (
                event.clientY -
                startY
            );


        applyTransform();

    }
);


map.addEventListener(
    "pointerup",
    function(event) {

        dragging =
            false;


        map.classList.remove(
            "dragging"
        );


        try {

            map.releasePointerCapture(
                event.pointerId
            );

        }
        catch (error) {

        }

    }
);


map.addEventListener(
    "pointercancel",
    function() {

        dragging =
            false;


        map.classList.remove(
            "dragging"
        );

    }
);


// ============================================================
// MOUSE WHEEL ZOOM
// ============================================================

map.addEventListener(
    "wheel",
    function(event) {

        event.preventDefault();


        const rect =
            map.getBoundingClientRect();


        const mouseX =
            event.clientX -
            rect.left;


        const mouseY =
            event.clientY -
            rect.top;


        const oldZoom =
            zoom;


        const factor =
            event.deltaY < 0
                ? 1.1
                : .9;


        zoom =
            Math.max(
                .2,
                Math.min(
                    3,
                    zoom * factor
                )
            );


        panX =
            mouseX -
            (
                mouseX -
                panX
            ) *
            (
                zoom /
                oldZoom
            );


        panY =
            mouseY -
            (
                mouseY -
                panY
            ) *
            (
                zoom /
                oldZoom
            );


        applyTransform();

    },
    {
        passive: false
    }
);


// ============================================================
// TOOLTIP
// ============================================================

function showTooltip(
    node,
    event
) {

    let content =
        "<div class='tooltip-title'>" +
        escapeHTML(
            node.label
        ) +
        "</div>";


    if (
        node.status
    ) {

        content +=
            "<div>" +
            escapeHTML(
                node.status
            ) +
            "</div>";

    }


    if (
        node.description
    ) {

        content +=
            "<div style='margin-top:8px'>" +
            escapeHTML(
                node.description
            ) +
            "</div>";

    }


    if (
        node.related_to
    ) {

        content +=
            "<div class='tooltip-related'>" +

            "<strong>Related to:</strong><br>" +

            escapeHTML(
                node.related_to
            );


        if (
            node.related_path
        ) {

            content +=
                "<br><small>" +
                escapeHTML(
                    node.related_path
                ) +
                "</small>";

        }


        content +=
            "</div>";

    }


    tooltip.innerHTML =
        content;


    tooltip.style.display =
        "block";


    moveTooltip(
        event
    );

}


function moveTooltip(
    event
) {

    let x =
        event.clientX +
        15;


    let y =
        event.clientY +
        15;


    const rect =
        tooltip.getBoundingClientRect();


    if (
        x + rect.width >
        window.innerWidth
    ) {

        x =
            event.clientX -
            rect.width -
            15;

    }


    if (
        y + rect.height >
        window.innerHeight
    ) {

        y =
            event.clientY -
            rect.height -
            15;

    }


    tooltip.style.left =
        x + "px";


    tooltip.style.top =
        y + "px";

}


function escapeHTML(
    value
) {

    const div =
        document.createElement(
            "div"
        );


    div.textContent =
        value;


    return div.innerHTML;

}


// ============================================================
// RESIZE
// ============================================================

window.addEventListener(
    "resize",
    function() {

        centre();

        render();

    }
);


// ============================================================
// INITIAL STATE
// ============================================================

rootOpen =
    false;


nodes.forEach(
    function(node) {

        open[node.id] =
            false;

    }
);


elements[0]
    .classList.add(
        "collapsed"
    );


centre();

render();


console.log(
    "================================"
);

console.log(
    "COLLABORATION MIND MAP READY"
);

console.log(
    "ROOT CLOSED"
);

console.log(
    "Click root -> Categories"
);

console.log(
    "Click category -> Subcategories"
);

console.log(
    "Click subcategory -> Stages"
);

console.log(
    "Click stage -> 3-column Modules"
);

console.log(
    "================================"
);


</script>

</body>

</html>
'''


# ============================================================
# INSERT DATA
# ============================================================

HTML = HTML.replace(
    "__NODES__",
    nodes_json
)

HTML = HTML.replace(
    "__LINKS__",
    links_json
)


# ============================================================
# WRITE HTML
# ============================================================

output_path = os.path.join(
    SOURCE_DIR,
    OUTPUT_FILE
)


with open(
    output_path,
    "w",
    encoding="utf-8"
) as file:

    file.write(
        HTML
    )


# ============================================================
# COMPLETE
# ============================================================

print()
print("=" * 70)
print("SUCCESS")
print("=" * 70)
print()
print(
    "HTML created:"
)
print(
    output_path
)
print()
print(
    "Excel files:",
    len(excel_files)
)
print(
    "Content items:",
    len(all_modules)
)
print(
    "Mind-map nodes:",
    len(nodes)
)
print()
print(
    "Interaction:"
)
print(
    "  Root       -> Categories"
)
print(
    "  Category   -> Subcategories"
)
print(
    "  Subcategory -> Stages"
)
print(
    "  Stage      -> 3-column Modules"
)
print()