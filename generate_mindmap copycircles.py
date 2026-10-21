import os
import glob
import json
import html
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


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def get_similarity(a, b):
    """
    Compare two pieces of text and return a similarity score
    between 0 and 1.
    """

    if not a or not b:
        return 0

    return SequenceMatcher(
        None,
        str(a),
        str(b)
    ).ratio()


def clean_text(value):
    """
    Convert Excel values safely into strings.
    """

    if value is None:
        return ""

    return str(value).strip()


# ============================================================
# FIND EXCEL FILES
# ============================================================

os.chdir(SOURCE_DIR)

excel_files = glob.glob("*.xlsx")

print(
    f"Found {len(excel_files)} Excel file(s)."
)


# ============================================================
# READ EXCEL DATA
# ============================================================

all_modules = []

for file in excel_files:

    print(
        f"Reading: {file}"
    )

    try:

        wb = openpyxl.load_workbook(
            file,
            data_only=True
        )

        sheet = wb.active

        # ----------------------------------------------------
        # Header / hierarchy information
        # ----------------------------------------------------

        d10 = clean_text(
            sheet["D10"].value
        ) or "Category"

        b10 = clean_text(
            sheet["B10"].value
        ) or "Sub-Category"

        f10 = clean_text(
            sheet["F10"].value
        ) or "Section"

        path_str = (
            f"{d10} > {b10} > {f10}"
        )

        # ----------------------------------------------------
        # Read content rows
        # ----------------------------------------------------

        for row in sheet.iter_rows(
            min_row=START_ROW,
            max_row=END_ROW,
            min_col=3,
            max_col=8
        ):

            # row indexes within C:H:
            #
            # row[0] = C
            # row[1] = D
            # row[2] = E
            # row[3] = F
            # row[4] = G
            # row[5] = H

            subtopic = clean_text(
                row[3].value
            )

            if not subtopic:
                continue

            status = clean_text(
                row[0].value
            ).lower()

            description = clean_text(
                row[5].value
            )

            all_modules.append({

                "d10": d10,

                "b10": b10,

                "f10": f10,

                "path": path_str,

                "status": status,

                "subtopic": subtopic,

                "description": description,

                "related_to": None,

                "related_path": None

            })

    except Exception as e:

        print(
            f"ERROR reading {file}: {e}"
        )


print(
    f"Loaded {len(all_modules)} content item(s)."
)


# ============================================================
# SIMILARITY CHECK
# ============================================================

print(
    "Checking for related content..."
)

for i, mod in enumerate(all_modules):

    for j, other in enumerate(all_modules):

        if i == j:
            continue

        if not mod["description"]:
            continue

        if not other["description"]:
            continue

        similarity = get_similarity(
            mod["description"],
            other["description"]
        )

        if similarity > SIMILARITY_THRESHOLD:

            mod["related_to"] = (
                other["subtopic"]
            )

            mod["related_path"] = (
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

for mod in all_modules:

    hierarchy[
        mod["d10"]
    ][
        mod["b10"]
    ][
        mod["f10"]
    ].append(mod)


# ============================================================
# MIND MAP DATA
# ============================================================

nodes = []
links = []

next_node_id = 0


def create_node(
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

    current_id = next_node_id

    next_node_id += 1

    return current_id


# ============================================================
# ROOT NODE
# ============================================================

root_id = create_node(
    "Collaboration Map",
    0
)


# ============================================================
# CREATE HIERARCHY NODES
# ============================================================

for d10, b10_dict in hierarchy.items():

    category_id = create_node(
        d10,
        1,
        root_id
    )

    for b10, f10_dict in b10_dict.items():

        subcategory_id = create_node(
            b10,
            2,
            category_id
        )

        for f10, modules in f10_dict.items():

            section_id = create_node(
                f10,
                3,
                subcategory_id
            )

            for mod in modules:

                create_node(

                    mod["subtopic"],

                    4,

                    section_id,

                    description=mod[
                        "description"
                    ],

                    status=mod[
                        "status"
                    ],

                    related_to=mod[
                        "related_to"
                    ],

                    related_path=mod[
                        "related_path"
                    ]

                )


print(
    f"Created {len(nodes)} mind-map nodes."
)


# ============================================================
# JSON DATA
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
#
# IMPORTANT:
#
# This is deliberately NOT an f-string.
#
# This prevents JavaScript {} from being interpreted
# by Python as f-string expressions.
# ============================================================

full_html = r"""<!DOCTYPE html>

<html lang="en">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>Collaboration Mind Map</title>


<style>

/* ============================================================
   GLOBAL
   ============================================================ */

* {
    box-sizing: border-box;
}

html,
body {

    width: 100%;
    height: 100%;

    margin: 0;
    padding: 0;

    overflow: hidden;

    font-family:
        Arial,
        Helvetica,
        sans-serif;

    background: #f4f7f9;
}


/* ============================================================
   HEADER
   ============================================================ */

.header {

    position: fixed;

    top: 0;
    left: 0;

    width: 100%;
    height: 65px;

    background: #005073;

    color: white;

    display: flex;

    align-items: center;

    justify-content: space-between;

    padding: 0 22px;

    z-index: 5000;

    box-shadow:
        0 2px 8px rgba(0, 0, 0, 0.25);

}

.header h1 {

    margin: 0;

    font-size: 23px;

    font-weight: 600;

}

.controls {

    display: flex;

    gap: 7px;

}

.controls button {

    border: none;

    background: white;

    color: #005073;

    padding: 8px 13px;

    border-radius: 5px;

    cursor: pointer;

    font-weight: bold;

}

.controls button:hover {

    background: #e5f1f5;

}


/* ============================================================
   MAP AREA
   ============================================================ */

#map-container {

    position: absolute;

    top: 65px;

    left: 0;

    right: 0;

    bottom: 0;

    overflow: hidden;

    cursor: grab;

    background:

        radial-gradient(
            circle at center,
            #ffffff 0%,
            #f4f7f9 70%
        );

}

#map-container.dragging {

    cursor: grabbing;

}


/* ============================================================
   SVG
   ============================================================ */

#connections {

    position: absolute;

    top: 0;
    left: 0;

    width: 100%;
    height: 100%;

    overflow: visible;

}


/* ============================================================
   CONNECTIONS
   ============================================================ */

.connection {

    fill: none;

    stroke-linecap: round;

    opacity: 0.65;

}

.connection.level1 {

    stroke: #005073;

    stroke-width: 5;

}

.connection.level2 {

    stroke: #3c8099;

    stroke-width: 3.5;

}

.connection.level3 {

    stroke: #86aab7;

    stroke-width: 2.5;

}

.connection.level4 {

    stroke: #b8c8ce;

    stroke-width: 2;

}


/* ============================================================
   NODE
   ============================================================ */

.node {

    position: absolute;

    transform:
        translate(-50%, -50%);

    border-radius: 50%;

    display: flex;

    align-items: center;

    justify-content: center;

    text-align: center;

    padding: 10px;

    cursor: pointer;

    user-select: none;

    overflow: hidden;

    word-break: break-word;

    transition:

        transform 0.18s ease,

        box-shadow 0.18s ease,

        opacity 0.2s ease;

    z-index: 10;

}


/* ============================================================
   NODE HOVER
   ============================================================ */

.node:hover {

    transform:
        translate(-50%, -50%)
        scale(1.08);

    box-shadow:
        0 8px 25px rgba(0, 0, 0, 0.25);

    z-index: 100;

}


/* ============================================================
   ROOT
   ============================================================ */

.level0 {

    width: 180px;
    height: 180px;

    background: #005073;

    color: white;

    border: 8px solid white;

    font-size: 22px;

    font-weight: bold;

    box-shadow:
        0 5px 25px rgba(0, 0, 0, 0.3);

    z-index: 1000;

}


/* ============================================================
   CATEGORY
   ============================================================ */

.level1 {

    width: 150px;
    height: 150px;

    background: #0076a8;

    color: white;

    border: 5px solid white;

    font-size: 16px;

    font-weight: bold;

}


/* ============================================================
   SUB-CATEGORY
   ============================================================ */

.level2 {

    width: 125px;
    height: 125px;

    background: #5f9db3;

    color: white;

    border: 4px solid white;

    font-size: 14px;

    font-weight: bold;

}


/* ============================================================
   SECTION
   ============================================================ */

.level3 {

    width: 110px;
    height: 110px;

    background: #d6e6eb;

    color: #003b50;

    border: 3px solid white;

    font-size: 12px;

    font-style: italic;

}


/* ============================================================
   CONTENT / TOPIC
   ============================================================ */

.level4 {

    width: 105px;
    height: 105px;

    background: white;

    color: #333;

    border: 4px solid #c5cdd1;

    font-size: 11px;

    box-shadow:
        0 3px 10px rgba(0, 0, 0, 0.12);

}


/* ============================================================
   STATUS
   ============================================================ */

.level4.mandatory {

    border-color: #d9534f;

}

.level4.optional {

    border-color: #5cb85c;

}

.level4.related {

    border-color: #6f42c1;

    background: #f8f5ff;

}


/* ============================================================
   HIDDEN
   ============================================================ */

.node.hidden {

    opacity: 0;

    pointer-events: none;

}

.connection.hidden {

    opacity: 0;

}


/* ============================================================
   TOOLTIP
   ============================================================ */

.tooltip {

    position: fixed;

    display: none;

    max-width: 380px;

    background:
        rgba(25, 25, 25, 0.96);

    color: white;

    padding: 15px;

    border-radius: 8px;

    font-size: 13px;

    line-height: 1.5;

    z-index: 10000;

    pointer-events: none;

    box-shadow:
        0 8px 30px rgba(0, 0, 0, 0.3);

}

.tooltip-title {

    font-size: 16px;

    font-weight: bold;

    margin-bottom: 8px;

}

.tooltip-status {

    display: inline-block;

    padding: 3px 7px;

    border-radius: 4px;

    background: #555;

    margin-bottom: 8px;

}

.tooltip-related {

    margin-top: 10px;

    padding-top: 10px;

    border-top:
        1px solid #666;

}


/* ============================================================
   LEGEND
   ============================================================ */

.legend {

    position: fixed;

    bottom: 18px;

    left: 18px;

    background: white;

    padding: 12px 15px;

    border-radius: 8px;

    box-shadow:
        0 3px 15px rgba(0, 0, 0, 0.15);

    z-index: 5000;

    font-size: 12px;

}

.legend-title {

    font-weight: bold;

    margin-bottom: 7px;

}

.legend-row {

    display: flex;

    align-items: center;

    margin: 5px 0;

}

.legend-dot {

    width: 13px;
    height: 13px;

    border-radius: 50%;

    margin-right: 7px;

}

.mandatory-dot {

    background: #d9534f;

}

.optional-dot {

    background: #5cb85c;

}

.related-dot {

    background: #6f42c1;

}


/* ============================================================
   INSTRUCTIONS
   ============================================================ */

.instructions {

    position: fixed;

    bottom: 18px;

    right: 18px;

    background:
        rgba(255, 255, 255, 0.95);

    padding: 10px 14px;

    border-radius: 7px;

    box-shadow:
        0 3px 12px rgba(0, 0, 0, 0.12);

    z-index: 5000;

    font-size: 11px;

    color: #555;

}


/* ============================================================
   RESPONSIVE
   ============================================================ */

@media (max-width: 900px) {

    .header h1 {

        font-size: 18px;

    }

    .level0 {

        width: 140px;
        height: 140px;

        font-size: 18px;

    }

    .level1 {

        width: 120px;
        height: 120px;

        font-size: 13px;

    }

    .level2 {

        width: 100px;
        height: 100px;

        font-size: 11px;

    }

    .level3,
    .level4 {

        width: 85px;
        height: 85px;

        font-size: 9px;

    }

    .instructions {

        display: none;

    }

}

</style>

</head>


<body>


<!-- ==========================================================
     HEADER
     ========================================================== -->

<div class="header">

    <h1>
        Collaboration Mind Map
    </h1>

    <div class="controls">

        <button onclick="resetMap()">
            Reset
        </button>

        <button onclick="zoomOut()">
            −
        </button>

        <button onclick="zoomIn()">
            +
        </button>

        <button onclick="fitMap()">
            Fit
        </button>

    </div>

</div>


<!-- ==========================================================
     MAP
     ========================================================== -->

<div id="map-container">

    <svg id="connections">

        <g id="connection-group"></g>

    </svg>

</div>


<!-- ==========================================================
     TOOLTIP
     ========================================================== -->

<div id="tooltip"
     class="tooltip">
</div>


<!-- ==========================================================
     LEGEND
     ========================================================== -->

<div class="legend">

    <div class="legend-title">
        Content Status
    </div>

    <div class="legend-row">

        <span class="legend-dot mandatory-dot">
        </span>

        Mandatory

    </div>

    <div class="legend-row">

        <span class="legend-dot optional-dot">
        </span>

        Optional

    </div>

    <div class="legend-row">

        <span class="legend-dot related-dot">
        </span>

        Related Content

    </div>

</div>


<!-- ==========================================================
     INSTRUCTIONS
     ========================================================== -->

<div class="instructions">

    Click a node to expand/collapse |
    Drag to move |
    Scroll to zoom |
    Hover over topics for details

</div>


<script>


// ============================================================
// DATA FROM PYTHON
// ============================================================

const nodes = __NODES_JSON__;

const links = __LINKS_JSON__;


// ============================================================
// MAP ELEMENTS
// ============================================================

const container =
    document.getElementById(
        "map-container"
    );

const connectionGroup =
    document.getElementById(
        "connection-group"
    );

const tooltip =
    document.getElementById(
        "tooltip"
    );


// ============================================================
// MAP STATE
// ============================================================

let scale = 1;

let offsetX = 0;

let offsetY = 0;

let dragging = false;

let dragStartX = 0;

let dragStartY = 0;


// ============================================================
// NODE LOOKUP
// ============================================================

const nodeById = {};

nodes.forEach(
    function(node) {

        nodeById[node.id] =
            node;

    }
);


// ============================================================
// CHILDREN LOOKUP
// ============================================================

const children = {};

nodes.forEach(
    function(node) {

        children[node.id] = [];

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
// EXPANSION STATE
// ============================================================

const expanded = {};

nodes.forEach(
    function(node) {

        expanded[node.id] = true;

    }
);


// ============================================================
// HTML NODE ELEMENTS
// ============================================================

const nodeElements = {};


// ============================================================
// SVG CONNECTION ELEMENTS
// ============================================================

const connectionElements = [];


// ============================================================
// CREATE NODES
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


        // ----------------------------------------------------
        // STATUS CLASSES
        // ----------------------------------------------------

        if (
            node.level === 4
        ) {

            const status =
                (
                    node.status || ""
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


        element.textContent =
            node.label;


        element.dataset.id =
            node.id;


        // ----------------------------------------------------
        // CLICK
        // ----------------------------------------------------

        element.addEventListener(
            "click",
            function(event) {

                event.stopPropagation();

                toggleNode(
                    node.id
                );

            }
        );


        // ----------------------------------------------------
        // HOVER
        // ----------------------------------------------------

        element.addEventListener(
            "mouseenter",
            function(event) {

                showTooltip(
                    node,
                    event
                );

            }
        );


        element.addEventListener(
            "mousemove",
            function(event) {

                moveTooltip(
                    event
                );

            }
        );


        element.addEventListener(
            "mouseleave",
            function() {

                hideTooltip();

            }
        );


        container.appendChild(
            element
        );


        nodeElements[
            node.id
        ] = element;

    }
);


// ============================================================
// CREATE CONNECTIONS
// ============================================================

links.forEach(
    function(link) {

        const path =
            document.createElementNS(
                "http://www.w3.org/2000/svg",
                "path"
            );

        const target =
            nodeById[
                link.target
            ];


        path.classList.add(
            "connection"
        );

        path.classList.add(
            "level" +
            target.level
        );


        connectionGroup.appendChild(
            path
        );


        connectionElements.push({

            link: link,

            element: path

        });

    }
);


// ============================================================
// LAYOUT SETTINGS
// ============================================================

const LEVEL_DISTANCE = {

    0: 0,

    1: 240,

    2: 210,

    3: 185,

    4: 160

};


// ============================================================
// BUILD TREE STRUCTURE
// ============================================================

function getChildren(id) {

    return children[id] || [];

}


function getVisibleDescendants(id) {

    let result = [];

    const childIds =
        getChildren(id);


    childIds.forEach(
        function(childId) {

            result.push(
                childId
            );


            if (
                expanded[id]
            ) {

                result =
                    result.concat(
                        getVisibleDescendants(
                            childId
                        )
                    );

            }

        }
    );


    return result;

}


// ============================================================
// CALCULATE BRANCH SIZE
//
// This determines how much vertical space each branch needs.
// ============================================================

function calculateBranchSize(
    nodeId
) {

    const childIds =
        getChildren(nodeId);


    if (
        childIds.length === 0
    ) {

        return 1;

    }


    let total = 0;


    childIds.forEach(
        function(childId) {

            if (
                expanded[nodeId]
            ) {

                total +=
                    calculateBranchSize(
                        childId
                    );

            }

        }
    );


    return Math.max(
        total,
        1
    );

}


// ============================================================
// LAYOUT A SINGLE BRANCH
// ============================================================

function layoutBranch(
    nodeId,
    direction,
    centerY,
    availableHeight,
    positions
) {

    const node =
        nodeById[nodeId];


    const childIds =
        getChildren(nodeId);


    if (
        childIds.length === 0 ||
        !expanded[nodeId]
    ) {

        return;

    }


    const branchSizes =
        childIds.map(
            function(childId) {

                return calculateBranchSize(
                    childId
                );

            }
        );


    const totalSize =
        branchSizes.reduce(
            function(sum, value) {

                return sum + value;

            },
            0
        );


    let currentY =
        centerY -
        (
            availableHeight *
            totalSize /
            2 /
            totalSize
        );


    const gap =
        Math.max(
            115,
            availableHeight /
            Math.max(
                totalSize,
                1
            )
        );


    childIds.forEach(
        function(childId, index) {

            const size =
                branchSizes[index];


            const childCenter =
                currentY +
                (
                    size *
                    gap /
                    2
                );


            const child =
                nodeById[
                    childId
                ];


            child.x =
                node.x +
                direction *
                LEVEL_DISTANCE[
                    child.level
                ];


            child.y =
                childCenter;


            currentY +=
                size * gap;


            layoutBranch(
                childId,
                direction,
                childCenter,
                gap * size,
                positions
            );

        }
    );

}


// ============================================================
// MAIN LAYOUT
// ============================================================

function calculateLayout() {

    const width =
        container.clientWidth;

    const height =
        container.clientHeight;


    const centerX =
        width / 2;

    const centerY =
        height / 2;


    const root =
        nodeById[0];


    root.x =
        centerX;

    root.y =
        centerY;


    const rootChildren =
        getChildren(0);


    const leftBranches = [];

    const rightBranches = [];


    rootChildren.forEach(
        function(id, index) {

            if (
                index %
                2 === 0
            ) {

                rightBranches.push(
                    id
                );

            } else {

                leftBranches.push(
                    id
                );

            }

        }
    );


    // --------------------------------------------------------
    // RIGHT SIDE
    // --------------------------------------------------------

    positionRootBranches(
        rightBranches,
        1,
        centerY,
        height
    );


    // --------------------------------------------------------
    // LEFT SIDE
    // --------------------------------------------------------

    positionRootBranches(
        leftBranches,
        -1,
        centerY,
        height
    );

}


// ============================================================
// POSITION ROOT BRANCHES
// ============================================================

function positionRootBranches(
    branchIds,
    direction,
    centerY,
    height
) {

    if (
        branchIds.length === 0
    ) {

        return;

    }


    let totalSize = 0;


    branchIds.forEach(
        function(id) {

            totalSize +=
                calculateBranchSize(
                    id
                );

        }
    );


    const verticalGap =
        Math.max(
            150,
            (
                height - 100
            ) /
            Math.max(
                totalSize,
                1
            )
        );


    let currentY =
        centerY -
        (
            totalSize *
            verticalGap /
            2
        );


    branchIds.forEach(
        function(id) {

            const branchSize =
                calculateBranchSize(
                    id
                );


            const branchCenter =
                currentY +
                (
                    branchSize *
                    verticalGap /
                    2
                );


            const node =
                nodeById[id];


            node.x =
                nodeById[0].x +
                direction *
                LEVEL_DISTANCE[1];


            node.y =
                branchCenter;


            currentY +=
                branchSize *
                verticalGap;


            layoutBranch(
                id,
                direction,
                branchCenter,
                branchSize *
                verticalGap,
                {}
            );

        }
    );

}


// ============================================================
// RENDER
// ============================================================

function render() {

    calculateLayout();


    // --------------------------------------------------------
    // Position nodes
    // --------------------------------------------------------

    nodes.forEach(
        function(node) {

            const element =
                nodeElements[
                    node.id
                ];


            const screenX =
                node.x *
                scale +
                offsetX;


            const screenY =
                node.y *
                scale +
                offsetY;


            element.style.left =
                screenX + "px";


            element.style.top =
                screenY + "px";


            element.style.transform =
                "translate(-50%, -50%) " +
                "scale(" +
                scale +
                ")";


            updateNodeVisibility(
                node
            );

        }
    );


    // --------------------------------------------------------
    // Draw connections
    // --------------------------------------------------------

    connectionElements.forEach(
        function(item) {

            const source =
                nodeById[
                    item.link.source
                ];

            const target =
                nodeById[
                    item.link.target
                ];


            const sourceX =
                source.x *
                scale +
                offsetX;


            const sourceY =
                source.y *
                scale +
                offsetY;


            const targetX =
                target.x *
                scale +
                offsetX;


            const targetY =
                target.y *
                scale +
                offsetY;


            const direction =
                target.x >= source.x
                    ? 1
                    : -1;


            const curve =
                Math.abs(
                    targetX -
                    sourceX
                ) * 0.45;


            const control1X =
                sourceX +
                direction *
                curve;


            const control2X =
                targetX -
                direction *
                curve;


            const pathData =
                "M " +
                sourceX +
                " " +
                sourceY +
                " C " +
                control1X +
                " " +
                sourceY +
                ", " +
                control2X +
                " " +
                targetY +
                ", " +
                targetX +
                " " +
                targetY;


            item.element.setAttribute(
                "d",
                pathData
            );


            const visible =
                isNodeVisible(
                    source.id
                ) &&
                isNodeVisible(
                    target.id
                );


            item.element.classList.toggle(
                "hidden",
                !visible
            );

        }
    );

}


// ============================================================
// VISIBILITY
// ============================================================

function isNodeVisible(id) {

    let current =
        nodeById[id];


    while (
        current.parent !== null
    ) {

        const parent =
            current.parent;


        if (
            !expanded[parent]
        ) {

            return false;

        }


        current =
            nodeById[parent];

    }


    return true;

}


function updateNodeVisibility(
    node
) {

    const element =
        nodeElements[
            node.id
        ];


    element.classList.toggle(
        "hidden",
        !isNodeVisible(
            node.id
        )
    );

}


// ============================================================
// TOGGLE
// ============================================================

function toggleNode(id) {

    if (
        !children[id] ||
        children[id].length === 0
    ) {

        return;

    }


    expanded[id] =
        !expanded[id];


    render();

}


// ============================================================
// TOOLTIP
// ============================================================

function showTooltip(
    node,
    event
) {

    let content =
        '<div class="tooltip-title">' +
        escapeHtml(
            node.label
        ) +
        "</div>";


    if (
        node.level === 4
    ) {

        if (
            node.status
        ) {

            content +=
                '<div class="tooltip-status">' +
                escapeHtml(
                    node.status
                ) +
                "</div>";

        }


        if (
            node.description
        ) {

            content +=
                "<div>" +
                escapeHtml(
                    node.description
                ) +
                "</div>";

        }


        if (
            node.related_to
        ) {

            content +=
                '<div class="tooltip-related">' +
                "<strong>Related to:</strong><br>" +
                escapeHtml(
                    node.related_to
                );


            if (
                node.related_path
            ) {

                content +=
                    "<br><small>" +
                    escapeHtml(
                        node.related_path
                    ) +
                    "</small>";

            }


            content +=
                "</div>";

        }

    } else {

        const childCount =
            children[node.id]
                ? children[node.id].length
                : 0;


        if (
            childCount > 0
        ) {

            content +=
                "<div>" +
                childCount +
                " child node" +
                (
                    childCount === 1
                        ? ""
                        : "s"
                ) +
                "</div>";

        }

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

    const padding = 15;

    let x =
        event.clientX +
        padding;

    let y =
        event.clientY +
        padding;


    const rect =
        tooltip.getBoundingClientRect();


    if (
        x + rect.width >
        window.innerWidth
    ) {

        x =
            event.clientX -
            rect.width -
            padding;

    }


    if (
        y + rect.height >
        window.innerHeight
    ) {

        y =
            event.clientY -
            rect.height -
            padding;

    }


    tooltip.style.left =
        x + "px";


    tooltip.style.top =
        y + "px";

}


function hideTooltip() {

    tooltip.style.display =
        "none";

}


// ============================================================
// ESCAPE HTML
// ============================================================

function escapeHtml(text) {

    const div =
        document.createElement(
            "div"
        );

    div.textContent =
        text;

    return div.innerHTML;

}


// ============================================================
// ZOOM
// ============================================================

function zoomIn() {

    scale *= 1.15;


    scale =
        Math.min(
            scale,
            2.5
        );


    render();

}


function zoomOut() {

    scale /= 1.15;


    scale =
        Math.max(
            scale,
            0.35
        );


    render();

}


// ============================================================
// RESET
// ============================================================

function resetMap() {

    scale = 1;

    offsetX = 0;

    offsetY = 0;


    nodes.forEach(
        function(node) {

            expanded[
                node.id
            ] = true;

        }
    );


    render();

}


// ============================================================
// FIT MAP
// ============================================================

function fitMap() {

    scale = 0.75;

    offsetX = 0;

    offsetY = 0;


    render();

}


// ============================================================
// DRAG / PAN
// ============================================================

container.addEventListener(
    "mousedown",
    function(event) {

        if (
            event.target.classList.contains(
                "node"
            )
        ) {

            return;

        }


        dragging = true;


        container.classList.add(
            "dragging"
        );


        dragStartX =
            event.clientX -
            offsetX;


        dragStartY =
            event.clientY -
            offsetY;

    }
);


container.addEventListener(
    "mousemove",
    function(event) {

        if (!dragging) {

            return;

        }


        offsetX =
            event.clientX -
            dragStartX;


        offsetY =
            event.clientY -
            dragStartY;


        render();

    }
);


container.addEventListener(
    "mouseup",
    function() {

        dragging = false;

        container.classList.remove(
            "dragging"
        );

    }
);


container.addEventListener(
    "mouseleave",
    function() {

        dragging = false;

        container.classList.remove(
            "dragging"
        );

    }
);


// ============================================================
// MOUSE WHEEL ZOOM
// ============================================================

container.addEventListener(
    "wheel",
    function(event) {

        event.preventDefault();


        if (
            event.deltaY < 0
        ) {

            scale *= 1.1;

        } else {

            scale /= 1.1;

        }


        scale =
            Math.max(
                0.35,
                Math.min(
                    scale,
                    2.5
                )
            );


        render();

    },
    {
        passive: false
    }
);


// ============================================================
// RESIZE
// ============================================================

window.addEventListener(
    "resize",
    function() {

        render();

    }
);


// ============================================================
// INITIAL RENDER
// ============================================================

render();


</script>

</body>

</html>
"""


# ============================================================
# INSERT PYTHON DATA INTO HTML
#
# We use replace() rather than an f-string.
# This is what prevents the JavaScript braces from causing
# Python syntax errors.
# ============================================================

full_html = full_html.replace(
    "__NODES_JSON__",
    nodes_json
)

full_html = full_html.replace(
    "__LINKS_JSON__",
    links_json
)


# ============================================================
# WRITE OUTPUT
# ============================================================

output_path = os.path.join(
    SOURCE_DIR,
    OUTPUT_FILE
)

with open(
    output_path,
    "w",
    encoding="utf-8"
) as output:

    output.write(
        full_html
    )


# ============================================================
# FINISHED
# ============================================================

print("")
print("=" * 60)
print("MIND MAP CREATED SUCCESSFULLY")
print("=" * 60)
print("")
print(
    f"Output: {output_path}"
)
print(
    f"Excel files: {len(excel_files)}"
)
print(
    f"Content items: {len(all_modules)}"
)
print(
    f"Mind-map nodes: {len(nodes)}"
)
print("")
print(
    "Open index.html in your browser."
)