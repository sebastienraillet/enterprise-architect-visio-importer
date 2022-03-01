from xmlrpc.client import Boolean
import win32com.client
from vsdx import VisioFile, Shape
from typing import Final, List

from os import listdir
from os.path import isfile, join

VISIO_INPUT_DIR: Final[str] = 'input_visio_file'

EA_ACTIVITY_DIAGRAM: Final[str] = 'Activity'

EA_ACTION_ELEMENT: Final[str] = 'Action'
EA_EVENT_ELEMENT: Final[str] = 'Event'
EA_ENTITY_ELEMENT: Final[str] = 'Entity'
EA_ACTIVITY_ELEMENT: Final[str] = 'Activity'
EA_CONSTRAINT_ELEMENT: Final[str] = 'Constraint'
EA_RISK_ELEMENT: Final[str] = 'Risk'
EA_OBJECT_ELEMENT: Final[str] = 'Object'

EA_CONTROL_FLOW_CONNECTOR: Final[str] = 'ControlFlow'

FULL_MODEL = 0


OLD_DOMAIN_EVENT_COLOR: Final[str] = '#f09609'
OLD_COMMAND_COLOR: Final[str] = '#1ba1e2'
OLD_USER_COLOR: Final[str] = '#e4ed3b'
OLD_POLICY_COLOR: Final[str] = '#7900bf'
OLD_EXTERNAL_SYSTEM_COLOR: Final[str] = '#f08dae'
OLD_VIEW_READ_MODEL_COLOR: Final[str] = '#75d175'
OLD_RISKS_COLOR: Final[str] = '#ff6556'
OLD_MEDIA_AGGREGATE: Final[str] = '#ffff00'


ALLOWED_DOMAIN_EVENT_COLOR: Final[str] = '#ffa95f'
ALLOWED_COMMAND_COLOR: Final[str] = '#a7cdf5'
ALLOWED_USER_COLOR: Final[str] = '#fff9b4'
ALLOWED_POLICY_COLOR: Final[str] = '#bd87c6'
ALLOWED_AGGREGATE_COLOR: Final[str] = '#f3d02b'
ALLOWED_EXTERNAL_SYSTEM_COLOR: Final[str] = '#eca1c4'
ALLOWED_VIEW_READ_MODEL_COLOR: Final[str] = '#d5f694'
ALLOWED_RISKS_COLOR: Final[str] = '#ee6d80'
ALLOWED_NONE_COLOR: Final[str] = None

ALLOWED_COLORS: List[str] = [ALLOWED_DOMAIN_EVENT_COLOR, ALLOWED_COMMAND_COLOR, ALLOWED_USER_COLOR,
                             ALLOWED_POLICY_COLOR, ALLOWED_AGGREGATE_COLOR, ALLOWED_EXTERNAL_SYSTEM_COLOR,
                             ALLOWED_VIEW_READ_MODEL_COLOR, ALLOWED_RISKS_COLOR]

OLD_COLORS: List[str] = [OLD_DOMAIN_EVENT_COLOR, OLD_COMMAND_COLOR, OLD_USER_COLOR, OLD_POLICY_COLOR,
                         OLD_EXTERNAL_SYSTEM_COLOR, OLD_VIEW_READ_MODEL_COLOR, OLD_RISKS_COLOR, OLD_MEDIA_AGGREGATE]

OLD_NEW_COLORS_MAPPING = {OLD_DOMAIN_EVENT_COLOR: ALLOWED_DOMAIN_EVENT_COLOR,
                          OLD_COMMAND_COLOR: ALLOWED_COMMAND_COLOR,
                          OLD_USER_COLOR: ALLOWED_USER_COLOR,
                          OLD_POLICY_COLOR: ALLOWED_POLICY_COLOR,
                          OLD_EXTERNAL_SYSTEM_COLOR: ALLOWED_EXTERNAL_SYSTEM_COLOR,
                          OLD_VIEW_READ_MODEL_COLOR: ALLOWED_VIEW_READ_MODEL_COLOR,
                          OLD_RISKS_COLOR: ALLOWED_RISKS_COLOR,
                          OLD_MEDIA_AGGREGATE: ALLOWED_AGGREGATE_COLOR}

COLOR_EA_ELEMENTS_MAPPING = {ALLOWED_DOMAIN_EVENT_COLOR: EA_ACTION_ELEMENT,
                             ALLOWED_COMMAND_COLOR: EA_ACTION_ELEMENT,
                             ALLOWED_USER_COLOR: EA_ENTITY_ELEMENT,
                             ALLOWED_POLICY_COLOR: EA_CONSTRAINT_ELEMENT,
                             ALLOWED_AGGREGATE_COLOR: EA_ACTIVITY_ELEMENT,
                             ALLOWED_EXTERNAL_SYSTEM_COLOR: EA_EVENT_ELEMENT,
                             ALLOWED_VIEW_READ_MODEL_COLOR: EA_OBJECT_ELEMENT,
                             ALLOWED_RISKS_COLOR: EA_RISK_ELEMENT,
                             ALLOWED_NONE_COLOR: EA_OBJECT_ELEMENT}

PIXEL_PER_INCHES: Final[int] = 96
PAGE_HEIGHT_INCHES: Final[int] = 11.70

VISIO_CONNECTORS = {}

class Connector:
    def __init__(self, p_visio_connector) -> None:
        self.m_visio_connector = p_visio_connector
        self.start_connector_side = None
        self.end_connector_side = None

    def is_valid(self) -> Boolean:
        return False if self.start_connector_side is None or self.end_connector_side is None else True


def is_connector(p_shape) -> bool:
    return True if p_shape.cell_value('ShapeRouteStyle') is not None else False


def fix_old_color(p_shape):
    l_current_color = p_shape.cell_value('FillForegnd')
    p_shape.set_cell_value('FillForegnd', OLD_NEW_COLORS_MAPPING[l_current_color])


def verify_shape_color(p_shape):
    l_sub_shapes = p_shape.sub_shapes()
    if l_sub_shapes:  # If there is sub shapes to check, go recursive call
        for l_shape in l_sub_shapes:
            if l_shape.shape_type is not None:
                verify_shape_color(l_shape)

    # Verify if the shape foreground exists, and it's color is allowed
    if p_shape.shape_type != 'Group':
        if p_shape.cell_value('FillForegnd') in OLD_COLORS:  # Old color present try to fix it
            fix_old_color(p_shape)
            print(f"The color for the element with the ID: {p_shape.ID} and the text: \"{p_shape.text[:-1]}\""
                  f" on page: \"{page.name}\" has been successfully replaced")
        if p_shape.cell_value('FillForegnd') not in ALLOWED_COLORS:
            print(f"The element with the ID: {p_shape.ID} and the text: \"{p_shape.text[:-1]}\""
                  f" on page: \"{page.name}\" is made of disallowed color")

def get_position(p_shape: Shape):
    x_position = 0
    y_position = 0

    if p_shape.parent.shape_type == 'Group':
        x_position, y_position = get_position(p_shape.parent)

    if p_shape.shape_type == 'Group':
        x_position += p_shape.x - (p_shape.width/2)
        y_position += p_shape.y - (p_shape.height/2)
    else:
        x_position += p_shape.x
        y_position += p_shape.y

    return (x_position, y_position)

def convert_shape_coordinates_to_EA(p_shape: Shape):
    # Variables until the next comments are in inch
    x_position, y_position = get_position(p_shape)
    x_left_top_corner = x_position - (p_shape.width/2)
    y_left_top_corner = PAGE_HEIGHT_INCHES - (y_position + (p_shape.height/2))
    
    # Convert inches in pixels
    width = (x_left_top_corner + p_shape.width)*PIXEL_PER_INCHES
    heigth = (y_left_top_corner + p_shape.height)*PIXEL_PER_INCHES

    x_left_top_corner *= PIXEL_PER_INCHES
    y_left_top_corner *= PIXEL_PER_INCHES

    return (x_left_top_corner, width, y_left_top_corner, heigth)


def convert_RGB_to_BGR(p_rgb_color: str):
    """Convert a RGB color to BGR color format 

        Keyword arguments:
        p_rgb_color -- The RGB color to convert in the form of "#ffaaff"
    """
    red = p_rgb_color[1:3]
    green = p_rgb_color[3:5]
    blue = p_rgb_color[5:7]
    return f"#{blue}{green}{red}"


def convert_RGB_to_EA_color(p_rgb_color: str):
    """Convert a RGB color to EA color format (decimal value of a color in BGR)

        Keyword arguments:
        p_rgb_color -- The RGB color to convert in the form of "#ffaaff"
    """

    l_bgr_color = convert_RGB_to_BGR(p_rgb_color) if p_rgb_color is not None else p_rgb_color
    l_EA_color = int(l_bgr_color[1:], 16) if l_bgr_color is not None else -1
    return l_EA_color


def store_connector(p_visio_connector: Shape, p_ea_element):
    # Store internally the connector to be able to build it later on when all the element
    # will be available inside the diagram.
    # This part also identify if the p_ea_element is the start side or the end side of
    # this connector
    if not p_visio_connector.ID in VISIO_CONNECTORS:
        l_connector = Connector(p_visio_connector)
        for connected_shape in p_visio_connector.connects:
            if connected_shape.shape.text.rstrip("\n") == p_ea_element.Name:
                if connected_shape.from_rel == 'BeginX':
                    l_connector.start_connector_side = p_ea_element.ElementGUID
                else:
                    l_connector.end_connector_side = p_ea_element.ElementGUID

        VISIO_CONNECTORS[p_visio_connector.ID] = l_connector
    else:
        l_connector = VISIO_CONNECTORS[p_visio_connector.ID]
        if l_connector.start_connector_side is None:
            l_connector.start_connector_side = p_ea_element.ElementGUID
        else:
            l_connector.end_connector_side = p_ea_element.ElementGUID

def create_EA_connectors(p_ea_repository):
    for connector in VISIO_CONNECTORS.values():
        if connector.is_valid():
            ea_element_start_connector_side = p_ea_repository.GetElementByGuid(connector.start_connector_side)
            l_connector = ea_element_start_connector_side.Connectors.AddNew("", EA_CONTROL_FLOW_CONNECTOR)
            l_connector.SupplierID = p_ea_repository.GetElementByGuid(connector.end_connector_side).ElementID
            l_connector.Update()

def convert_shape_to_EA_element(p_shape: Shape, p_use_case_package, p_use_case_diagram):
    l_sub_shapes = p_shape.sub_shapes()
    if l_sub_shapes:  # If there is sub shapes to check, go recursive call
        for l_shape in l_sub_shapes:
            if l_shape.shape_type is not None:
                convert_shape_to_EA_element(l_shape, p_use_case_package, p_use_case_diagram)

    if p_shape.shape_type != 'Group':
        l_rgb_color = p_shape.cell_value('FillForegnd')
        l_object_type = COLOR_EA_ELEMENTS_MAPPING[l_rgb_color]

        l_element = p_use_case_package.Elements.AddNew(p_shape.text.rstrip("\n"), l_object_type)
        l_element.SetAppearance(1, 0, convert_RGB_to_EA_color(l_rgb_color))
        l_element.Update()
        
        for connector in p_shape.connected_shapes:
            store_connector(connector, l_element)
        
        x_left_top_corner, width, y_left_top_corner, heigth = convert_shape_coordinates_to_EA(p_shape)
        l_position = f"l={x_left_top_corner};r={width};t={y_left_top_corner};b={heigth};"

        l_diagram_object = p_use_case_diagram.DiagramObjects.AddNew(l_position, "")
        l_diagram_object.ElementID = l_element.ElementID
        l_diagram_object.Update()

if __name__ == "__main__":
    try:
        eaApp = win32com.client.Dispatch("EA.App")
    except:
        print(f"Impossible to connect to EA, please verify Enterprise Architect is open")
        exit()
    mEaRep = eaApp.Repository
    if mEaRep.ConnectionString == '':
        print(f"EA has no Model loaded")
        exit()
    else:
        print(f"Connecting...{mEaRep.ConnectionString}")

    l_visio_files = [f for f in listdir(VISIO_INPUT_DIR) if isfile(join(VISIO_INPUT_DIR, f))]

    mEaRep.BatchAppend = True
    mEaRep.EnableUIUpdates = False
    for visio_file_path in l_visio_files:
        # For each file, we create a new package inside EA
        l_root_package = mEaRep.Models.GetAt(0).Packages.AddNew(visio_file_path, "")
        l_root_package.Update()

        with VisioFile(join(VISIO_INPUT_DIR, visio_file_path)) as vis:
            for page in vis.pages:
                # For each pages inside the Visio file, we create an activity diagram
                # with the page name
                l_diagram = l_root_package.Diagrams.AddNew(page.name, EA_ACTIVITY_DIAGRAM)
                l_diagram.Update()

                # Iterate over each shape of this page
                shapes = page.sub_shapes()
                for shape in shapes:
                    if not is_connector(shape):
                        verify_shape_color(shape)
                        convert_shape_to_EA_element(shape, l_root_package, l_diagram)

                create_EA_connectors(mEaRep)
                VISIO_CONNECTORS = {}
    
    mEaRep.RefreshModelView(FULL_MODEL)
    mEaRep.BatchAppend = False
    mEaRep.EnableUIUpdates = True
