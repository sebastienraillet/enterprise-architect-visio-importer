import win32com.client
import argparse
import pathlib
import xlsxwriter

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
EA_STATE_ELEMENT: Final[str] = 'State'
EA_TEXT_ELEMENT: Final[str] = 'Text'

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
                             ALLOWED_VIEW_READ_MODEL_COLOR, ALLOWED_RISKS_COLOR, ALLOWED_NONE_COLOR]

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
                             ALLOWED_VIEW_READ_MODEL_COLOR: EA_STATE_ELEMENT,
                             ALLOWED_RISKS_COLOR: EA_RISK_ELEMENT,
                             ALLOWED_NONE_COLOR: EA_TEXT_ELEMENT}

PIXEL_PER_INCHES: Final[int] = 96
PAGE_HEIGHT_INCHES: Final[int] = 11.70

VISIO_CONNECTORS = {}

class Connector:
    def __init__(self, p_visio_connector) -> None:
        self.m_visio_connector = p_visio_connector
        self.start_connector_side = None
        self.end_connector_side = None

    def is_valid(self) -> bool:
        return False if self.start_connector_side is None or self.end_connector_side is None else True

class VisioShape:
    def __init__(self, p_internal_shape: Shape) -> None:
        self.m_internal_visio_shape = p_internal_shape
        self.page = None

    @property
    def ID(self) -> str:
        return self.m_internal_visio_shape.ID

    @property
    def text(self) -> str:
        return self.m_internal_visio_shape.text[:-1]

    @property
    def color(self):
        return self.m_internal_visio_shape.cell_value('FillForegnd')

    @color.setter
    def color(self, p_color: str):
        self.m_internal_visio_shape.set_cell_value('FillForegnd', p_color)

    @property
    def shape_type(self) -> str:
        return self.m_internal_visio_shape.shape_type
    
    @property
    def parent(self):
        return VisioShape(self.m_internal_visio_shape.parent)

    @property
    def x(self) -> float:
        return self.m_internal_visio_shape.x

    @property
    def y(self) -> float:
        return self.m_internal_visio_shape.y

    @property
    def width(self) -> float:
        return self.m_internal_visio_shape.width

    @property
    def height(self) -> float:
        return self.m_internal_visio_shape.height

    def get_position(self):
        x_position = 0
        y_position = 0

        if self.parent.shape_type == 'Group':
            x_position, y_position = self.parent.get_position()

        if self.shape_type == 'Group':
            x_position += self.shape.x - (self.width/2)
            y_position += self.y - (self.height/2)
        else:
            x_position += self.x
            y_position += self.y

        return (x_position, y_position)

    def fix_old_color(self):
        # Old color present try to fix it
        if self.color in OLD_COLORS:  
            self.color = OLD_NEW_COLORS_MAPPING[self.color]

    def is_color_allowed(self) -> bool:
        return True if self.color in ALLOWED_COLORS else False


class VisioPage:
    def __init__(self, p_name: str) -> None:
        self.m_shapes = []
        self.name = p_name

    def add_shape(self, p_shape: VisioShape):
        for sub_shape in p_shape.m_internal_visio_shape.sub_shapes():
            if sub_shape.shape_type is not None:
                l_visio_shape = VisioShape(sub_shape)
                self.add_shape(l_visio_shape)
                
        if p_shape.m_internal_visio_shape.shape_type != 'Group':
            if not p_shape in self.m_shapes:
                p_shape.page = self
                self.m_shapes.append(p_shape)
            else:
                raise ValueError(f"shape: {p_shape.text} already exist in page {self.name}")

    @property
    def shapes(self):
        return self.m_shapes


class VisioFileToImport:
    def __init__(self, p_path: pathlib.Path) -> None:
        self.m_pages = {}
        self.name = p_path.stem
        self.path = p_path

    def add_page(self, p_page: VisioPage):
        if not p_page.name in self.m_pages.keys():
            self.m_pages[p_page.name] = p_page
        else:
            raise KeyError(f"The page: {p_page.name} already exist in the file: {self.m_name}")
    
    @property
    def pages(self):
        return self.m_pages.values()




def is_connector(p_shape) -> bool:
    return True if p_shape.cell_value('ShapeRouteStyle') is not None else False


def convert_shape_coordinates_to_EA(p_shape: VisioShape):
    # Variables until the next comments are in inch
    x_position, y_position = p_shape.get_position()
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

def convert_shape_to_EA_element(p_shape: VisioShape, p_use_case_package, p_use_case_diagram):
    l_rgb_color = p_shape.color
    l_object_type = COLOR_EA_ELEMENTS_MAPPING[l_rgb_color]

    l_element = p_use_case_package.Elements.AddNew(p_shape.text.rstrip("\n"), l_object_type)
    if l_object_type == EA_TEXT_ELEMENT:
        # To have information display on a text element, note should be put on it
        l_element.Notes = p_shape.text
    l_element.SetAppearance(1, 0, convert_RGB_to_EA_color(l_rgb_color))
    l_element.Update()
    
    for connector in p_shape.m_internal_visio_shape.connected_shapes:
        store_connector(connector, l_element)
    
    x_left_top_corner, width, y_left_top_corner, heigth = convert_shape_coordinates_to_EA(p_shape)
    l_position = f"l={x_left_top_corner};r={width};t={y_left_top_corner};b={heigth};"

    l_diagram_object = p_use_case_diagram.DiagramObjects.AddNew(l_position, "")
    l_diagram_object.ElementID = l_element.ElementID
    l_diagram_object.Update()

def generate_color_report(p_visio_file: VisioFileToImport, p_bad_shape_list):
    l_excel_file_path = str(p_visio_file.path).replace(".vsdx", ".xlsx")
    l_workbook = xlsxwriter.Workbook(l_excel_file_path)
    l_worksheet = l_workbook.add_worksheet()

    # Build header row
    l_bold = l_workbook.add_format({'bold': True})
    l_worksheet.write(0, 0, "Page name", l_bold)
    l_worksheet.write(0, 1, "Shape ID", l_bold)
    l_worksheet.write(0, 2, "Text", l_bold)
    l_worksheet.write(0, 3, "Disallowed color", l_bold)

    l_row_index = 1
    for bad_shape in p_bad_shape_list:
        l_worksheet.write(l_row_index, 0, bad_shape.page.name)
        l_worksheet.write(l_row_index, 1, bad_shape.ID)
        l_worksheet.write(l_row_index, 2, bad_shape.text)
        l_worksheet.write(l_row_index, 3, bad_shape.color)
        l_row_index += 1

    l_workbook.close()     


def build_files_list_to_import(p_path):
    l_visio_files_to_import = []
    if p_path.exists():
        if p_path.is_file(): 
            if p_path.suffix == ".vsdx":
                # User provided only one file to import and it's a vsdx file
                l_visio_files_to_import.append(p_path)
            else:
                print(f"The path {p_path} isn't a path to a Visio (.vsdx) file")
        elif p_path.is_dir():
            # User provided a directory inside we should look for all vsdx files
            l_visio_files_to_import = list(p_path.glob('*.vsdx'))

            # If the list is empty, we found no vsdx files in the directory. We inform
            # the user about that
            if not l_visio_files_to_import:
                print(f"The directory you provided doesn't contain any Visio (*.vsdx) files")
        else:
            print(f"The path you provided isn't a path to a file or a directory")
    else:
        print(f"The specified path {p_path} doesn't exist")

    return l_visio_files_to_import

if __name__ == "__main__":
    l_parser = argparse.ArgumentParser(prog="Enterprise Architect Visio Event storming importer")
    l_parser.add_argument("Path", help="Path to a Visio file (or a directory containing multiple Visio files) to be imported in EA",
                          type=pathlib.Path)
    l_parser.add_argument("GUID", help="Enterprise Architect GUID on which the imported elements will be added")
    l_parser.add_argument("-c", "--check-colors-only", help="Verify only if colors used in Visio diagram are \
                          compliant with Event storming without doing any addition to Enterprise Architect", action="store_true")
    l_parser.add_argument("-g", "--generate-color-report", help="Generate an Excel report for each visio file imported listing all the colors \
                          which aren't compliant", action="store_true")
    l_parser.add_argument("--fix-colors", help="Try to fix the colors used in the Visio diagram if they aren't \
                          compliant with Event storming convention", action="store_true")
    l_parser.add_argument("--dry-run", help="run the script without doing the import in Enterprise Architect", action="store_true")
    args = l_parser.parse_args()

    l_visio_file_to_work_on = []
    l_visio_files_path = build_files_list_to_import(args.Path)
    if not l_visio_files_path:
        print(f"No visio file to import, exiting...")
        exit()

    # Store all the data we need to work on
    for visio_file_path in l_visio_files_path:
        with VisioFile(str(visio_file_path)) as vis:
            l_visio_file_to_import = VisioFileToImport(visio_file_path)
            for page in vis.pages:
                l_visio_page_to_import = VisioPage(page.name)
                 # Iterate over each shape of this page
                shapes = page.sub_shapes()
                for shape in shapes:
                    if not is_connector(shape):
                        l_visio_shape = VisioShape(shape)
                        l_visio_page_to_import.add_shape(l_visio_shape)
                l_visio_file_to_import.add_page(l_visio_page_to_import)
            l_visio_file_to_work_on.append(l_visio_file_to_import)

    l_shape_bad_color = []
    l_bad_color_found = False
    # First check if colors used in Visio are the correct one (based on Event Storming template).
    # Then, generates a report if the user askeded to do it
    for visio_file in l_visio_file_to_work_on:
        for page in visio_file.pages:
            # Iterate over each shape of this page
            for shape in page.shapes:
                shape.fix_old_color()
                if not shape.is_color_allowed():
                    l_shape_bad_color.append(shape)
                    # print(f"The element with the ID: {shape.ID} and the text: \"{shape.text}\""
                    #       f" on page: \"{page.name}\" is made of a disallowed color: {shape.color}")
        if l_shape_bad_color:
            l_bad_color_found = True
            if args.generate_color_report:
                generate_color_report(visio_file, l_shape_bad_color)
            l_shape_bad_color = []

    # Then, we add the element in Enterprise architect if all the colors check are ok and
    # the user asked to do it
    if (not args.check_colors_only) and (not args.dry_run) and (not l_bad_color_found):
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

        # Get the node from which we will start to add the imported elements from Visio. This node
        # is recovered through the GUID provided by the user
        l_root_node = mEaRep.GetPackageByGuid(args.GUID)

        mEaRep.BatchAppend = True
        mEaRep.EnableUIUpdates = False
        for visio_file in l_visio_file_to_work_on:
            # For each file, we create a new package inside EA
            l_root_package = l_root_node.Packages.AddNew(visio_file.name, "")
            l_root_package.Update()

            for page in visio_file.pages:
                # For each pages inside the Visio file, we create an activity diagram
                # with the page name
                l_diagram = l_root_package.Diagrams.AddNew(page.name, EA_ACTIVITY_DIAGRAM)
                l_diagram.Update()

                # Iterate over each shape of this page
                for shape in page.shapes:
                    convert_shape_to_EA_element(shape, l_root_package, l_diagram)

                create_EA_connectors(mEaRep)
                VISIO_CONNECTORS = {}
        
        mEaRep.RefreshModelView(FULL_MODEL)
        mEaRep.BatchAppend = False
        mEaRep.EnableUIUpdates = True
