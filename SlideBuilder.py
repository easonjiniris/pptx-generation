"""
SlideBuilder class - Auto-generated from slide inspection.
Each method corresponds to a slide template in the library.
"""

import os
from typing import Dict, List
import win32com.client


class SlideBuilder:
    VISIBILITY = 1
    _powerpoint_instance = None

    def __init__(self, library_path: str, output_path: str):
        """
        Initialize SlideBuilder with library and output paths.

        Args:
            library_path: Path to the source presentation library
            output_path: Path to the destination presentation
        """
        self.library_path = os.path.abspath(library_path)
        self.output_path = os.path.abspath(output_path)
        self.presentation = None
        self.slide = None
        self._shape_cache = {}

    @classmethod
    def get_powerpoint_instance(cls):
        """Get or create a PowerPoint application instance."""
        if cls._powerpoint_instance is None:
            cls._powerpoint_instance = win32com.client.Dispatch(
                "PowerPoint.Application")
            cls._powerpoint_instance.Visible = cls.VISIBILITY
        return cls._powerpoint_instance

    @classmethod
    def quit_powerpoint(cls):
        """Close the PowerPoint application."""
        if cls._powerpoint_instance is not None:
            cls._powerpoint_instance.Quit()
            cls._powerpoint_instance = None

    def create_blank(self, template_path: str = None) -> None:
        """
        Create a new blank presentation, optionally based on a template.

        Args:
            template_path: Optional path to a template presentation.
                          If None, uses the library_path as template.
        """
        powerpoint = self.get_powerpoint_instance()

        template = template_path if template_path else self.library_path

        if template:
            # Create from template
            presentation = powerpoint.Presentations.Open(
                template, ReadOnly=True)
            # Delete all slides
            while presentation.Slides.Count > 0:
                presentation.Slides(1).Delete()
        else:
            # Create completely blank
            presentation = powerpoint.Presentations.Add()

        presentation.SaveAs(self.output_path)
        presentation.Close()

    def open_output(self):
        """Open the output presentation for editing."""
        if self.presentation is None:
            powerpoint = self.get_powerpoint_instance()
            self.presentation = powerpoint.Presentations.Open(self.output_path)

    def close_output(self):
        """Close the output presentation."""
        if self.presentation is not None:
            self.presentation.Close()
            self.presentation = None

    def save_output(self):
        """Save the output presentation."""
        if self.presentation is not None:
            self.presentation.Save()

    def copy_slide(self, slide_index: int) -> int:
        """
        Copy a single slide from the library presentation to the output presentation.

        Args:
            slide_index: 0-based index of the slide to copy from the library

        Returns:
            The 1-based index of the newly inserted slide in the output presentation
        """
        powerpoint = self.get_powerpoint_instance()

        library_pres = None

        try:
            library_pres = powerpoint.Presentations.Open(
                self.library_path, ReadOnly=True)
            self.open_output()

            slide_to_copy = library_pres.Slides(slide_index + 1)
            slide_to_copy.Copy()

            new_slide_index = self.presentation.Slides.Count + 1
            self.presentation.Slides.Paste(new_slide_index)

            self.save_output()

            return new_slide_index
        finally:
            if library_pres:
                library_pres.Close()

    def copy_slides(self, slide_indices: List[int]) -> List[int]:
        """
        Copy multiple slides from the library presentation to the output presentation.

        Args:
            slide_indices: List of 0-based indices of slides to copy from the library

        Returns:
            List of 1-based indices of the newly inserted slides in the output presentation
        """
        powerpoint = self.get_powerpoint_instance()

        library_pres = None

        try:
            library_pres = powerpoint.Presentations.Open(
                self.library_path, ReadOnly=True)
            self.open_output()

            new_slide_indices = []

            for slide_index in slide_indices:
                slide_to_copy = library_pres.Slides(slide_index + 1)
                slide_to_copy.Copy()

                new_slide_index = self.presentation.Slides.Count + 1
                self.presentation.Slides.Paste(new_slide_index)
                new_slide_indices.append(new_slide_index)

            self.save_output()

            return new_slide_indices
        finally:
            if library_pres:
                library_pres.Close()

    def _build_shape_cache(self):
        """Build a cache of shapes indexed by name."""
        if self.slide is None:
            return
        self._shape_cache = {}
        for shape in self.slide.Shapes:
            self._shape_cache[shape.Name.lower()] = shape

    def _get_shape(self, name: str):
        """Get shape by name (case-insensitive)."""
        return self._shape_cache.get(name.lower())

    def _set_text(self, shape_name: str, text: str):
        """Set text content for a shape."""
        shape = self._get_shape(shape_name)
        if shape and hasattr(shape, "HasTextFrame") and shape.HasTextFrame:
            shape.TextFrame.TextRange.Text = text

    def _set_bullets(self, shape_name: str, items: list):
        """Set bullet points for a shape."""
        shape = self._get_shape(shape_name)
        if shape and hasattr(shape, "HasTextFrame") and shape.HasTextFrame:
            text = "\n".join(str(item) for item in items)
            shape.TextFrame.TextRange.Text = text
            # Apply bullet formatting
            try:
                for paragraph in shape.TextFrame.TextRange.Paragraphs():
                    paragraph.ParagraphFormat.Bullet.Visible = True
            except:
                pass

    def _set_table_cell(self, shape_name: str, row: int, col: int, text: str):
        """Set text content for a specific table cell.

        Args:
            shape_name: Name of the table shape
            row: Row number (1-indexed, 1 is the top row)
            col: Column number (1-indexed, 1 is the leftmost column)
            text: Text to set in the cell
        """
        shape = self._get_shape(shape_name)
        if shape and hasattr(shape, "HasTable") and shape.HasTable:
            try:
                shape.Table.Cell(
                    row, col).Shape.TextFrame.TextRange.Text = text
            except:
                pass

    def _set_group_text(self, group_name: str, shape_name: str, text: str):
        """Set text content for a shape within a group.

        Args:
            group_name: Name of the group shape
            shape_name: Name of the shape within the group
            text: Text to set in the shape
        """
        group = self._get_shape(group_name)
        if group and hasattr(group, "GroupItems"):
            try:
                for item in group.GroupItems:
                    if item.Name.lower() == shape_name.lower():
                        if hasattr(item, "HasTextFrame") and item.HasTextFrame:
                            item.TextFrame.TextRange.Text = text
                        break
            except:
                pass

    def fill_slide_type_title(self, title: str):
        self.slide = self.presentation.Slides(1)
        self._build_shape_cache()

        self._set_text("Title 1", title)

    def fill_slide_type_4(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "source" in content:
            self._set_text("ZoneTexte 8", content["source"])

        if "description" in content:
            self._set_text("ZoneTexte 19", content["description"])

    def fill_slide_type_5(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "source" in content:
            self._set_text("ZoneTexte 8", content["source"])

        if "description" in content:
            self._set_text("ZoneTexte 19", content["description"])

    def fill_slide_type_6(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "source" in content:
            self._set_text("source", content["source"])

        if "description" in content:
            self._set_text("Rectangle 26", content["description"])

    def fill_slide_type_7(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

    def fill_slide_type_8(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "summary_title_1" in content:
            self._set_text("ZoneTexte 4", content["summary_title_1"])

        if "summary_title_2" in content:
            self._set_text("ZoneTexte 6", content["summary_title_2"])

        if "zonetexte_8" in content:
            self._set_text("ZoneTexte 8", content["zonetexte_8"])

        if "source_1" in content:
            self._set_text("ZoneTexte 6", content["source_1"])

        if "source_2" in content:
            self._set_text("ZoneTexte 10", content["source_2"])

    def fill_slide_type_9(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "description" in content:
            self._set_text("Content Placeholder 3", content["description"])

        if "source" in content:
            self._set_text("ZoneTexte 10", content["source"])

    def fill_slide_type_10(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "description" in content:
            self._set_text("Content Placeholder 3", content["description"])

        if "graph_title" in content:
            self._set_text("TextBox 2", content["graph_title"])

    def fill_slide_type_11(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "cause_1" in content:
            self._set_text("Rectangle 7", content["cause_1"])

        if "conclusion_1" in content:
            self._set_text("Rectangle 8", content["conclusion_1"])

        if "consequence_1" in content:
            self._set_text("ZoneTexte 32", content["consequence_1"])

        if "cause_2" in content:
            self._set_text("Rectangle 13", content["cause_2"])

        if "conclusion_2" in content:
            self._set_text("Rectangle 10", content["conclusion_2"])

        if "consequence_2" in content:
            self._set_text("ZoneTexte 32", content["consequence_2"])

        if "cause_3" in content:
            self._set_text("Rectangle 14", content["cause_3"])

        if "conclusion_3" in content:
            self._set_text("Rectangle 12", content["conclusion_3"])

        if "consequence_3" in content:
            self._set_text("ZoneTexte 32", content["consequence_3"])

        if "benefit" in content:
            self._set_text("Ellipse 17", content["benefit"])

    def fill_slide_type_12(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "criteria_1" in content:
            self._set_table_cell("Table 7", 1, 1, content["criteria_1"])

        if "description_1" in content:
            self._set_table_cell("Table 7", 1, 2, content["description_1"])

        if "criteria_2" in content:
            self._set_table_cell("Table 7", 2, 1, content["criteria_2"])

        if "description_2" in content:
            self._set_table_cell("Table 7", 2, 2, content["description_2"])

        if "criteria_3" in content:
            self._set_table_cell("Table 7", 3, 1, content["criteria_3"])

        if "description_3" in content:
            self._set_table_cell("Table 7", 3, 2, content["description_3"])

        if "criteria_4" in content:
            self._set_table_cell("Table 7", 4, 1, content["criteria_4"])

        if "description_4" in content:
            self._set_table_cell("Table 7", 4, 2, content["description_4"])

        if "key_consequences" in content:
            self._set_text("Google Shape;1738;p230",
                           content["key_consequences"])

    def fill_slide_type_13(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "description" in content:
            self._set_text("ZoneTexte 18", content["description"])

        if "key_messages" in content:
            self._set_text("Google Shape;1738;p230", content["key_messages"])

    def fill_slide_type_15(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 7", title)

        if "company_description" in content:
            self._set_text("Rectangle 5", content["company_description"])

        if "our_role" in content:
            self._set_text("Rectangle 15", content["our_role"])

        if "results" in content:
            self._set_text("Rectangle 12", content["results"])

        if "quote" in content:
            self._set_text("Rectangle 18", content["quote"])

    def fill_slide_type_16(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "reference_1" in content:
            self._set_text("Rectangle 29", content["reference_1"])

        if "info_1" in content:
            self._set_text("Rectangle 28", content["info_1"])

        if "reference_2" in content:
            self._set_text("Rectangle 32", content["reference_2"])

        if "info_2" in content:
            self._set_text("Rectangle 31", content["info_2"])

    def fill_slide_type_17(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "reference_1" in content:
            self._set_text("Rectangle 29", content["reference_1"])

        if "info_1" in content:
            self._set_text("Rectangle 28", content["info_1"])

        if "reference_2" in content:
            self._set_text("Rectangle 18", content["reference_2"])

        if "info_2" in content:
            self._set_text("Rectangle 17", content["info_2"])

        if "reference_3" in content:
            self._set_text("Rectangle 24", content["reference_3"])

        if "info_3" in content:
            self._set_text("Rectangle 23", content["info_3"])

    def fill_slide_type_18(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "reference_1" in content:
            self._set_text("Rectangle 42", content["reference_1"])

        if "info_1" in content:
            self._set_text("Rectangle 43", content["info_1"])

        if "reference_2" in content:
            self._set_text("Rectangle 28", content["reference_2"])

        if "info_2" in content:
            self._set_text("Rectangle 31", content["info_2"])

        if "reference_3" in content:
            self._set_text("Rectangle 34", content["reference_3"])

        if "info_3" in content:
            self._set_text("Rectangle 35", content["info_3"])

        if "reference_4" in content:
            self._set_text("Rectangle 38", content["reference_4"])

        if "info_4" in content:
            self._set_text("Rectangle 39", content["info_4"])

    def fill_slide_type_19(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "roles_and_expertise" in content:
            self._set_text("Rectangle 7", content["roles_and_expertise"])

        if "education_and_languages" in content:
            self._set_text("Rectangle 4", content["education_and_languages"])

        if "selected_experience" in content:
            self._set_text("Rectangle 8", content["selected_experience"])

    def fill_slide_type_20(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "key_message" in content:
            self._set_text("Content Placeholder 12", content["key_message"])

        if "name_1" in content:
            self._set_text("Rectangle 16", content["name_1"])

        if "role_1" in content:
            self._set_text("Rectangle 19", content["role_1"])

        if "name_2" in content:
            self._set_text("Rectangle 16-2", content["name_2"])

        if "role_2" in content:
            self._set_text("Rectangle 19-2", content["role_2"])

        if "name_3" in content:
            self._set_text("Rectangle 16-3", content["name_3"])

        if "role_3" in content:
            self._set_text("Rectangle 19-3", content["role_3"])

        if "name_4" in content:
            self._set_text("Rectangle 11", content["name_4"])

        if "role_4" in content:
            self._set_text("Rectangle 19-4", content["role_4"])

    def fill_slide_type_21(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "key_message" in content:
            self._set_text("Content Placeholder 12", content["key_message"])

        if "name_1" in content:
            self._set_text("Rectangle 16", content["name_1"])

        if "role_1" in content:
            self._set_text("Rectangle 19", content["role_1"])

        if "name_2" in content:
            self._set_text("Rectangle 16-2", content["name_2"])

        if "role_2" in content:
            self._set_text("Rectangle 19-2", content["role_2"])

        if "name_3" in content:
            self._set_text("Rectangle 16-3", content["name_3"])

        if "role_3" in content:
            self._set_text("Rectangle 19-3", content["role_3"])

        if "name_4" in content:
            self._set_text("Rectangle 11", content["name_4"])

        if "role_4" in content:
            self._set_text("Rectangle 19-4", content["role_4"])

        if "name_5" in content:
            self._set_text("Rectangle 16-5", content["name_5"])

        if "role_5" in content:
            self._set_text("Rectangle 19-5", content["rol5_4"])

        if "name_6" in content:
            self._set_text("Rectangle 18", content["name_6"])

        if "role_6" in content:
            self._set_text("Rectangle 19-6", content["role_6"])

    def fill_slide_type_22(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "name_1" in content:
            self._set_text("Rectangle 11", content["name_1"])
        info_1 = ""
        if "title_1" in content:
            info_1 += content["title_1"]
        if "role_1" in content:
            info_1 += content["role_1"]
        if "expertise_1" in content:
            info_1 += content["expertise_1"]
        self._set_text("Rectangle 19", info_1)

        if "name_2" in content:
            self._set_text("Rectangle 30", content["name_2"])
        info_2 = ""
        if "title_2" in content:
            info_2 += content["title_2"]
        if "role_2" in content:
            info_2 += content["role_2"]
        if "expertise_2" in content:
            info_2 += content["expertise_2"]
        self._set_text("Rectangle 19-2", info_2)

        if "name_3" in content:
            self._set_text("Rectangle 31", content["name_3"])
        info_3 = ""
        if "title_3" in content:
            info_3 += content["title_3"]
        if "role_3" in content:
            info_3 += content["role_3"]
        if "expertise_3" in content:
            info_3 += content["expertise_3"]
        self._set_text("Rectangle 19-3", info_3)

        if "name_4" in content:
            self._set_text("Rectangle 32", content["name_4"])
        info_4 = ""
        if "title4" in content:
            info_4 += content["title_4"]
        if "role_4" in content:
            info_4 += content["role_4"]
        if "expertise_4" in content:
            info_4 += content["expertise_4"]
        self._set_text("Rectangle 20", info_4)

    def fill_slide_type_24(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "objective_1" in content:
            self._set_table_cell("Table 26", 1, 2, content["objective_1"])

        if "objective_2" in content:
            self._set_table_cell("Table 26", 2, 2, content["objective_2"])

        if "objective_3" in content:
            self._set_table_cell("Table 26", 3, 2, content["objective_3"])

        if "objective_4" in content:
            self._set_table_cell("Table 26", 4, 2, content["objective_4"])

    def fill_slide_type_25(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "principle_1" in content:
            self._set_table_cell("Espace réservé du contenu 4",
                                 1, 1, content["principle_1"])

        if "context_1" in content:
            self._set_table_cell(
                "Espace réservé du contenu 4", 1, 2, content["context_1"])

        if "principle_2" in content:
            self._set_table_cell("Espace réservé du contenu 4",
                                 2, 1, content["principle_2"])

        if "context_2" in content:
            self._set_table_cell(
                "Espace réservé du contenu 4", 2, 2, content["context_2"])

        if "principle_3" in content:
            self._set_table_cell("Espace réservé du contenu 4",
                                 3, 1, content["principle_3"])

        if "context_3" in content:
            self._set_table_cell(
                "Espace réservé du contenu 4", 3, 2, content["context_3"])

    def fill_slide_type_47(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "idea_1" in content:
            self._set_text("Rectangle 31", content["idea_1"])

        if "idea_2" in content:
            self._set_text("Rectangle 32", content["idea_2"])

    def fill_slide_type_48(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "idea_1" in content:
            self._set_text("Rectangle 31", content["idea_1"])

        if "idea_2" in content:
            self._set_text("Rectangle 32", content["idea_2"])

        if "idea_3" in content:
            self._set_text("Rectangle 33", content["idea_3"])

    def fill_slide_type_49(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "idea_1" in content:
            self._set_text("Rectangle 31", content["idea_1"])

        if "idea_2" in content:
            self._set_text("Rectangle 32", content["idea_2"])

        if "idea_3" in content:
            self._set_text("Rectangle 33", content["idea_3"])

        if "idea_4" in content:
            self._set_text("Rectangle 14", content["idea_4"])

    def fill_slide_type_50(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "idea_1" in content:
            self._set_text("Forme libre : forme 6", content["idea_1"])

        if "idea_2" in content:
            self._set_text("Forme libre : forme 9", content["idea_2"])

        if "idea_3" in content:
            self._set_text("Forme libre : forme 11", content["idea_3"])

    def fill_slide_type_51(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "idea_1" in content:
            self._set_text("Forme libre : forme 6", content["idea_1"])

        if "idea_2" in content:
            self._set_text("Forme libre : forme 9", content["idea_2"])

        if "idea_3" in content:
            self._set_text("Forme libre : forme 11", content["idea_3"])

        if "idea_4" in content:
            self._set_text("Forme libre : forme 18", content["idea_4"])

    def fill_slide_type_52(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "idea_1" in content:
            self._set_text("Forme libre : forme 33", content["idea_1"])

        if "idea_2" in content:
            self._set_text("Forme libre : forme 34", content["idea_2"])

        if "idea_3" in content:
            self._set_text("Forme libre : forme 35", content["idea_3"])

        if "idea_4" in content:
            self._set_text("Forme libre : forme 42", content["idea_4"])

        if "idea_5" in content:
            self._set_text("Forme libre : forme 36", content["idea_5"])

    def fill_slide_type_53(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 4", title)

        if "topic_1" in content:
            self._set_text("ZoneTexte 22", content["topic_1"])

        if "detail_1" in content:
            self._set_text("ZoneTexte 58", content["detail_1"])

        if "topic_2" in content:
            self._set_text("ZoneTexte 51", content["topic_2"])

        if "detail_2" in content:
            self._set_text("ZoneTexte 16", content["detail_2"])

        if "topic_3" in content:
            self._set_text("ZoneTexte 52", content["topic_3"])

        if "detail_3" in content:
            self._set_text("ZoneTexte 17", content["detail_3"])

    def fill_slide_type_54(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 4", title)

        if "topic_1" in content:
            self._set_text("ZoneTexte 22", content["topic_1"])

        if "detail_1" in content:
            self._set_text("ZoneTexte 58", content["detail_1"])

        if "topic_2" in content:
            self._set_text("ZoneTexte 51", content["topic_2"])

        if "detail_2" in content:
            self._set_text("ZoneTexte 18", content["detail_2"])

        if "topic_3" in content:
            self._set_text("ZoneTexte 52", content["topic_3"])

        if "detail_3" in content:
            self._set_text("ZoneTexte 23", content["detail_3"])

        if "topic_4" in content:
            self._set_text("ZoneTexte 53", content["topic_4"])

        if "detail_4" in content:
            self._set_text("ZoneTexte 24", content["detail_4"])

    def fill_slide_type_55(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 5", title)

        if "topic_1" in content:
            self._set_text("Rectangle 28", content["topic_1"])

        if "detail_1" in content:
            self._set_text("Rectangle 29", content["detail_1"])

        if "topic_2" in content:
            self._set_text("Rectangle 30", content["topic_2"])

        if "detail_2" in content:
            self._set_text("Rectangle 51", content["detail_2"])

    def fill_slide_type_56(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 5", title)

        if "topic_1" in content:
            self._set_text("Rectangle 28", content["topic_1"])

        if "detail_1" in content:
            self._set_text("Rectangle 29", content["detail_1"])

        if "topic_2" in content:
            self._set_text("Rectangle 30", content["topic_2"])

        if "detail_2" in content:
            self._set_text("Rectangle 51", content["detail_2"])

        if "topic_3" in content:
            self._set_text("Rectangle 31", content["topic_3"])

        if "detail_3" in content:
            self._set_text("Rectangle 53", content["detail_3"])

    def fill_slide_type_57(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 5", title)

        if "topic_1" in content:
            self._set_text("Rectangle 28", content["topic_1"])

        if "detail_1" in content:
            self._set_text("Rectangle 29", content["detail_1"])

        if "topic_2" in content:
            self._set_text("Rectangle 54", content["topic_2"])

        if "detail_2" in content:
            self._set_text("Rectangle 50", content["detail_2"])

        if "topic_3" in content:
            self._set_text("Rectangle 59", content["topic_3"])

        if "detail_3" in content:
            self._set_text("Rectangle 57", content["detail_3"])

        if "topic_4" in content:
            self._set_text("Rectangle 64", content["topic_4"])

        if "detail_4" in content:
            self._set_text("Rectangle 62", content["detail_4"])

    def fill_slide_type_58(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        msg = ""
        if "idea_1" in content:
            msg += content["idea_1"]
            msg += "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 37", msg)

        msg = ""
        if "idea_2" in content:
            msg += content["idea_2"]
            msg += "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 39", msg)

        msg = ""
        if "idea_3" in content:
            msg += content["idea_3"]
            msg += "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 38", msg)

    def fill_slide_type_59(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        msg = ""
        if "idea_1" in content:
            msg += content["idea_1"]
            msg += "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 12", msg)

        msg = ""
        if "idea_2" in content:
            msg += content["idea_2"]
            msg += "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 14", msg)

        msg = ""
        if "idea_3" in content:
            msg += content["idea_3"]
            msg += "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 13", msg)

        msg = ""
        if "idea_4" in content:
            msg += content["idea_4"]
            msg += "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 40", msg)

    def fill_slide_type_60(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        msg = ""
        if "idea_1" in content:
            msg += content["idea_1"]
            msg += "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 37", msg)

        msg = ""
        if "idea_2" in content:
            msg += content["idea_2"]
            msg += "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 22", msg)

        msg = ""
        if "idea_3" in content:
            msg += content["idea_3"]
            msg += "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 26", msg)

        msg = ""
        if "idea_4" in content:
            msg += content["idea_4"]
            msg += "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 38", msg)

        msg = ""
        if "idea_5" in content:
            msg += content["idea_5"]
            msg += "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("Rectangle 23", msg)

    def fill_slide_type_61(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        msg = ""
        if "idea_1" in content:
            msg += content["idea_1"]
            msg += "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 16", msg)

        msg = ""
        if "idea_2" in content:
            msg += content["idea_2"]
            msg += "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 18", msg)

        msg = ""
        if "idea_3" in content:
            msg += content["idea_3"]
            msg += "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 28", msg)

        msg = ""
        if "idea_4" in content:
            msg += content["idea_4"]
            msg += "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 17", msg)

        msg = ""
        if "idea_5" in content:
            msg += content["idea_5"]
            msg += "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("Rectangle 19", msg)

        msg = ""
        if "idea_6" in content:
            msg += content["idea_6"]
            msg += "\n"
        if "detail_6" in content:
            msg += content["detail_6"]
        self._set_text("Rectangle 29", msg)

    def fill_slide_type_62(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "idea_1" in content:
            self._set_text("ZoneTexte 5", content["idea_1"])

        if "detail_1" in content:
            self._set_text("ZoneTexte 6", content["detail_1"])

        if "idea_2" in content:
            self._set_text("ZoneTexte 7", content["idea_2"])

        if "detail_2" in content:
            self._set_text("ZoneTexte 9", content["detail_2"])

    def fill_slide_type_63(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "idea_1" in content:
            self._set_text("ZoneTexte 5", content["idea_1"])

        if "detail_1" in content:
            self._set_text("ZoneTexte 6", content["detail_1"])

        if "idea_2" in content:
            self._set_text("ZoneTexte 7", content["idea_2"])

        if "detail_2" in content:
            self._set_text("ZoneTexte 9", content["detail_2"])

        if "idea_3" in content:
            self._set_text("ZoneTexte 15", content["idea_3"])

        if "detail_3" in content:
            self._set_text("ZoneTexte 16", content["detail_3"])

    def fill_slide_type_64(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "criteria_1" in content:
            self._set_text("TextBox 53", content["criteria_1"])

        if "criteria_2" in content:
            self._set_text("TextBox 54", content["criteria_2"])

        if "quadrant_1" in content:
            self._set_text("Freeform 5", content["quadrant_1"])

        if "description_1" in content:
            self._set_text("TextBox 19", content["description_1"])

        if "quadrant_2" in content:
            self._set_text("Freeform 6", content["quadrant_2"])

        if "description_2" in content:
            self._set_text("TextBox 24", content["description_2"])

        if "quadrant_3" in content:
            self._set_text("Freeform 8", content["quadrant_3"])

        if "description_3" in content:
            self._set_text("TextBox 42", content["description_3"])

        if "quadrant_4" in content:
            self._set_text("Freeform 7", content["quadrant_4"])

        if "description_4" in content:
            self._set_text("TextBox 46", content["description_4"])

    def fill_slide_type_65(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "stake" in content:
            self._set_text("Oval 40", content["stake"])

        if "quadrant_1" in content:
            self._set_table_cell("Tableau 46", 1, 1, content["quadrant_1"])

        if "description_1" in content:
            self._set_table_cell("Tableau 46", 2, 1, content["description_1"])

        if "quadrant_2" in content:
            self._set_table_cell("Tableau 46-2", 1, 1, content["quadrant_2"])

        if "description_2" in content:
            self._set_table_cell("Tableau 46-2", 2, 1,
                                 content["description_2"])

        if "quadrant_3" in content:
            self._set_table_cell("Tableau 46-3", 1, 1, content["quadrant_3"])

        if "description_3" in content:
            self._set_table_cell("Tableau 46-3", 2, 1,
                                 content["description_3"])

        if "quadrant_4" in content:
            self._set_table_cell("Tableau 46-4", 1, 1, content["quadrant_4"])

        if "description_4" in content:
            self._set_table_cell("Tableau 46-4", 2, 1,
                                 content["description_4"])

    def fill_slide_type_66(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "cause_1" in content:
            self._set_text("Content Placeholder 4-1c", content["cause_1"])

        if "result_1" in content:
            self._set_text("Content Placeholder 4-1r", content["result_1"])

        if "cause_2" in content:
            self._set_text("Content Placeholder 4-2c", content["cause_1"])

        if "result_2" in content:
            self._set_text("Content Placeholder 4-2r", content["result_2"])

        if "cause_3" in content:
            self._set_text("Content Placeholder 4-3c", content["cause_3"])

        if "result_3" in content:
            self._set_text("Content Placeholder 4-3r", content["result_3"])

        if "cause_4" in content:
            self._set_text("Content Placeholder 4-4c", content["cause_4"])

        if "result_4" in content:
            self._set_text("Content Placeholder 4-4r", content["result_4"])

    def fill_slide_type_67(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "cause_1" in content:
            self._set_text("Content Placeholder 4-1c", content["cause_1"])

        if "result_1" in content:
            self._set_text("Content Placeholder 4-1r", content["result_1"])

        if "cause_2" in content:
            self._set_text("Content Placeholder 4-2c", content["cause_1"])

        if "result_2" in content:
            self._set_text("Content Placeholder 4-2r", content["result_2"])

        if "cause_3" in content:
            self._set_text("Content Placeholder 4-3c", content["cause_3"])

        if "result_3" in content:
            self._set_text("Content Placeholder 4-3r", content["result_3"])

    def fill_slide_type_68(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "cause_1" in content:
            self._set_text("Content Placeholder 4-1c", content["cause_1"])

        if "result_1" in content:
            self._set_text("Content Placeholder 4-1r", content["result_1"])

        if "cause_2" in content:
            self._set_text("Content Placeholder 4-2c", content["cause_1"])

        if "result_2" in content:
            self._set_text("Content Placeholder 4-2r", content["result_2"])

    def fill_slide_type_69(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 5", title)

        if "pro" in content:
            self._set_text("ZoneTexte 71", content["pro"])

        if "con" in content:
            self._set_text("ZoneTexte 72", content["con"])

        zones = ["ZoneTexte 77", "ZoneTexte 78",
                 "ZoneTexte 80", "ZoneTexte 86", "ZoneTexte 86-2"]
        if "detail_1" in content:
            points = content["detail_1"].splitlines()
            for i in range(len(points)):
                self._set_text(zones[i], points[i])

        zones = ["ZoneTexte 47", "ZoneTexte 48",
                 "ZoneTexte 49", "ZoneTexte 50", "ZoneTexte 50-2"]
        if "detail_2" in content:
            points = content["detail_2"].splitlines()
            for i in range(len(points)):
                self._set_text(zones[i], points[i])

    def fill_slide_type_71(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "idea_1" in content:
            self._set_text("Freeform 17", content["idea_1"])

        if "idea_2" in content:
            self._set_text("Freeform 12", content["idea_2"])

        if "idea_3" in content:
            self._set_text("Freeform 8", content["idea_3"])

        if "idea_4" in content:
            self._set_text("Freeform 13", content["idea_4"])

        if "idea_5" in content:
            self._set_text("Freeform 15", content["idea_5"])

        if "idea_6" in content:
            self._set_text("Freeform 14", content["idea_6"])

        if "idea_7" in content:
            self._set_text("Freeform 16", content["idea_7"])

        if "idea_8" in content:
            self._set_text("Freeform 11", content["idea_8"])

        if "idea_9" in content:
            self._set_text("Freeform 10", content["idea_9"])

        if "idea_10" in content:
            self._set_text("Freeform 18", content["idea_10"])

    def fill_slide_type_72(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "title" in content:
            self._set_text("Freeform 14", content["title"])

        if "key_1" in content:
            self._set_text("Freeform 15", content["key_1"])

        if "detail_1" in content:
            self._set_text("Espace réservé du contenu 5", content["detail_1"])

        if "key_2" in content:
            self._set_text("Freeform 16", content["key_2"])

        if "detail_2" in content:
            self._set_text("Espace réservé du contenu 5-2",
                           content["detail_2"])

        if "key_3" in content:
            self._set_text("Freeform 10", content["key_3"])

        if "detail_3" in content:
            self._set_text("Espace réservé du contenu 5-3",
                           content["detail_3"])

        if "key_4" in content:
            self._set_text("Freeform 18", content["key_4"])

        if "detail_4" in content:
            self._set_text("Espace réservé du contenu 5-4",
                           content["detail_4"])

        if "key_5" in content:
            self._set_text("Freeform 8", content["key_5"])

        if "detail_5" in content:
            self._set_text("Espace réservé du contenu 5-5",
                           content["detail_5"])

        if "key_6" in content:
            self._set_text("Freeform 13", content["key_6"])

        if "detail_6" in content:
            self._set_text("Espace réservé du contenu 5-6",
                           content["detail_6"])

    def fill_slide_type_73(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "idea_1" in content:
            self._set_text("ZoneTexte 15", content["idea_1"])

        if "idea_2" in content:
            self._set_text("ZoneTexte 632", content["idea_2"])

        if "idea_3" in content:
            self._set_text("ZoneTexte 634", content["idea_3"])

        if "idea_4" in content:
            self._set_text("ZoneTexte 635", content["idea_4"])

        if "idea_5" in content:
            self._set_text("ZoneTexte 633", content["idea_5"])

    def fill_slide_type_74(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "idea_1" in content:
            self._set_group_text(
                "Groupe 115", "ZoneTexte 109", content["idea_1"])

        if "detail_1" in content:
            self._set_text("TextBox 91", content["detail_1"])

        if "idea_2" in content:
            self._set_group_text(
                "Groupe 115", "ZoneTexte 111", content["idea_2"])

        if "detail_2" in content:
            self._set_text("TextBox 91-2", content["detail_2"])

        if "idea_3" in content:
            self._set_group_text(
                "Groupe 115", "ZoneTexte 112", content["idea_3"])

        if "detail_3" in content:
            self._set_text("TextBox 91-3", content["detail_3"])

        if "idea_4" in content:
            self._set_group_text(
                "Groupe 115", "ZoneTexte 113", content["idea_4"])

        if "detail_4" in content:
            self._set_text("TextBox 91-4", content["detail_4"])

    def fill_slide_type_75(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "objective_1" in content:
            self._set_text("Rectangle 25", content["objective_1"])

        if "objective_2" in content:
            self._set_text("Rectangle 26", content["objective_2"])

        if "objective_3" in content:
            self._set_text("Rectangle 27", content["objective_3"])

    def fill_slide_type_76(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "objective_1" in content:
            self._set_text("Rectangle 14", content["objective_1"])

        if "objective_2" in content:
            self._set_text("Rectangle 15", content["objective_2"])

        if "objective_3" in content:
            self._set_text("Rectangle 16", content["objective_3"])

        if "objective_4" in content:
            self._set_text("Rectangle 17", content["objective_4"])

    def fill_slide_type_77(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "key_1" in content:
            self._set_table_cell("Table 7", 1, 1, content["key_1"])

        if "detail_1" in content:
            self._set_table_cell("Table 7", 1, 2, content["detail_1"])

        if "key_2" in content:
            self._set_table_cell("Table 7", 2, 1, content["key_2"])

        if "detail_2" in content:
            self._set_table_cell("Table 7", 2, 2, content["detail_2"])

        if "key_3" in content:
            self._set_table_cell("Table 7", 3, 1, content["key_3"])

        if "detail_3" in content:
            self._set_table_cell("Table 7", 3, 2, content["detail_3"])

        if "key_4" in content:
            self._set_table_cell("Table 7", 4, 1, content["key_4"])

        if "detail_4" in content:
            self._set_table_cell("Table 7", 4, 2, content["detail_4"])

        msg = ""
        if "summary" in content:
            msg += "Summary:\n" + content["summary"] + "\n"
        if "next_steps" in content:
            msg += "Next Steps:\n" + content["next_steps"] + "\n"
        self._set_table_cell("Table 7", 1, 3, msg)

    def fill_slide_type_78(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "objective" in content:
            self._set_text("Heptagone 3", content["objective"])

        msg = ""
        if "criteria_1" in content:
            msg += content["criteria_1"] + "\n"
        if "description_1" in content:
            msg += content["description_1"]
        self._set_text("Espace réservé du contenu 5", msg)

        msg = ""
        if "criteria_2" in content:
            msg += content["criteria_2"] + "\n"
        if "description_2" in content:
            msg += content["description_2"]
        self._set_text("Espace réservé du contenu 5-2", msg)

        msg = ""
        if "criteria_3" in content:
            msg += content["criteria_3"] + "\n"
        if "description_3" in content:
            msg += content["description_3"]
        self._set_text("Espace réservé du contenu 5-3", msg)

        msg = ""
        if "criteria_4" in content:
            msg += content["criteria_4"] + "\n"
        if "description_4" in content:
            msg += content["description_4"]
        self._set_text("Espace réservé du contenu 5-4", msg)

        msg = ""
        if "criteria_5" in content:
            msg += content["criteria_5"] + "\n"
        if "description_5" in content:
            msg += content["description_5"]
        self._set_text("Espace réservé du contenu 5-5", msg)

        msg = ""
        if "criteria_6" in content:
            msg += content["criteria_6"] + "\n"
        if "description_6" in content:
            msg += content["description_6"]
        self._set_text("Espace réservé du contenu 5-6", msg)

        msg = ""
        if "criteria_7" in content:
            msg += content["criteria_7"] + "\n"
        if "description_7" in content:
            msg += content["description_7"]
        self._set_text("Espace réservé du contenu 5-7", msg)

    def fill_slide_type_79(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "objective" in content:
            self._set_text("Ellipse 2", content["objective"])

        msg = ""
        if "criteria_1" in content:
            msg += content["criteria_1"] + "\n"
        if "description_1" in content:
            msg += content["description_1"]
        self._set_text("Espace réservé du contenu 5", msg)

        msg = ""
        if "criteria_2" in content:
            msg += content["criteria_2"] + "\n"
        if "description_2" in content:
            msg += content["description_2"]
        self._set_text("Espace réservé du contenu 5-2", msg)

        msg = ""
        if "criteria_3" in content:
            msg += content["criteria_3"] + "\n"
        if "description_3" in content:
            msg += content["description_3"]
        self._set_text("Espace réservé du contenu 5-3", msg)

        msg = ""
        if "criteria_4" in content:
            msg += content["criteria_4"] + "\n"
        if "description_4" in content:
            msg += content["description_4"]
        self._set_text("Espace réservé du contenu 5-4", msg)

        msg = ""
        if "criteria_5" in content:
            msg += content["criteria_5"] + "\n"
        if "description_5" in content:
            msg += content["description_5"]
        self._set_text("Espace réservé du contenu 5-5", msg)

        msg = ""
        if "criteria_6" in content:
            msg += content["criteria_6"] + "\n"
        if "description_6" in content:
            msg += content["description_6"]
        self._set_text("Espace réservé du contenu 5-6", msg)

        msg = ""
        if "criteria_7" in content:
            msg += content["criteria_7"] + "\n"
        if "description_7" in content:
            msg += content["description_7"]
        self._set_text("Espace réservé du contenu 5-7", msg)

    def fill_slide_type_80(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        msg = ""
        if "principle_1" in content:
            msg += content["principle_1"] + "\n"
        if "description_1" in content:
            msg += content["description_1"]
        self._set_text("Rectangle 12", msg)

        msg = ""
        if "principle_2" in content:
            msg += content["principle_2"] + "\n"
        if "description_2" in content:
            msg += content["description_2"]
        self._set_text("Rectangle 13", msg)

        msg = ""
        if "principle_3" in content:
            msg += content["principle_3"] + "\n"
        if "description_3" in content:
            msg += content["description_3"]
        self._set_text("Rectangle 14", msg)

        msg = ""
        if "principle_4" in content:
            msg += content["principle_4"] + "\n"
        if "description_4" in content:
            msg += content["description_4"]
        self._set_text("Rectangle 15", msg)

    def fill_slide_type_81(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        msg = ""
        if "step_1" in content:
            msg += content["step_1"] + "\n"
        if "description_1" in content:
            msg += content["description_1"]
        self._set_text("Rectangle 35", msg)

        msg = ""
        if "step_2" in content:
            msg += content["step_2"] + "\n"
        if "description_2" in content:
            msg += content["description_2"]
        self._set_text("Rectangle 36", msg)

        msg = ""
        if "step_3" in content:
            msg += content["step_3"] + "\n"
        if "description_3" in content:
            msg += content["description_3"]
        self._set_text("Rectangle 37", msg)

        msg = ""
        if "step_4" in content:
            msg += content["step_4"] + "\n"
        if "description_4" in content:
            msg += content["description_4"]
        self._set_text("Rectangle 38", msg)

        msg = ""
        if "step_4" in content:
            msg += content["step_4"] + "\n"
        if "description_4" in content:
            msg += content["description_4"]
        self._set_text("Rectangle 39", msg)

    def fill_slide_type_83(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "scenario_1" in content:
            self._set_text("Rectangle 7", content["scenario_1"])

        if "detail_1" in content:
            self._set_text("TextBox 13", content["detail_1"])

        if "scenario_2" in content:
            self._set_text("Rectangle 8", content["scenario_2"])

        if "detail_2" in content:
            self._set_text("TextBox 13-2", content["detail_2"])

    def fill_slide_type_84(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        msg = ""
        if "scenario_1" in content:
            msg += content["scenario_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 3", msg)

        msg = ""
        if "scenario_2" in content:
            msg += content["scenario_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 4", msg)

    def fill_slide_type_85(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "scenario_1" in content:
            self._set_text("Rectangle 7", content["scenario_1"])

        if "detail_1" in content:
            self._set_text("TextBox 13", content["detail_1"])

        if "scenario_2" in content:
            self._set_text("Rectangle 8", content["scenario_2"])

        if "detail_2" in content:
            self._set_text("TextBox 13-2", content["detail_2"])

        if "scenario_3" in content:
            self._set_text("Rectangle 9", content["scenario_3"])

        if "detail_3" in content:
            self._set_text("TextBox 13-3", content["detail_3"])

    def fill_slide_type_86(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        msg = ""
        if "scenario_1" in content:
            msg += content["scenario_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 3", msg)

        msg = ""
        if "scenario_2" in content:
            msg += content["scenario_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 4", msg)

        msg = ""
        if "scenario_3" in content:
            msg += content["scenario_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 5", msg)

    def fill_slide_type_87(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "scenario_1" in content:
            self._set_text("Rectangle 7", content["scenario_1"])

        if "detail_1" in content:
            self._set_text("TextBox 13", content["detail_1"])

        if "scenario_2" in content:
            self._set_text("Rectangle 8", content["scenario_2"])

        if "detail_2" in content:
            self._set_text("TextBox 13-2", content["detail_2"])

        if "scenario_3" in content:
            self._set_text("Rectangle 9", content["scenario_3"])

        if "detail_3" in content:
            self._set_text("TextBox 13-3", content["detail_3"])

        if "scenario_4" in content:
            self._set_text("Rectangle 17", content["scenario_4"])

        if "detail_4" in content:
            self._set_text("TextBox 13-4", content["detail_4"])

    def fill_slide_type_88(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        msg = ""
        if "scenario_1" in content:
            msg += content["scenario_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 3", msg)

        msg = ""
        if "scenario_2" in content:
            msg += content["scenario_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 4", msg)

        msg = ""
        if "scenario_3" in content:
            msg += content["scenario_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 5", msg)

        msg = ""
        if "scenario_4" in content:
            msg += content["scenario_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 16", msg)

    def fill_slide_type_89(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 8", title)

        msg = ""
        if "scenario_1" in content:
            msg += content["scenario_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 3", msg)

        msg = ""
        if "scenario_2" in content:
            msg += content["scenario_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 4", msg)

        msg = ""
        if "scenario_3" in content:
            msg += content["scenario_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 5", msg)

        msg = ""
        if "scenario_4" in content:
            msg += content["scenario_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 6", msg)

        msg = ""
        if "scenario_5" in content:
            msg += content["scenario_5"] + "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("Rectangle 34", msg)

    def fill_slide_type_90(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 8", title)

        msg = ""
        if "scenario_1" in content:
            msg += content["scenario_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 3", msg)

        msg = ""
        if "scenario_2" in content:
            msg += content["scenario_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 4", msg)

        msg = ""
        if "scenario_3" in content:
            msg += content["scenario_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 5", msg)

        msg = ""
        if "scenario_4" in content:
            msg += content["scenario_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 6", msg)

        msg = ""
        if "scenario_5" in content:
            msg += content["scenario_5"] + "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("Rectangle 34", msg)

    def fill_slide_type_91(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 6", title)

        msg = ""
        if "scenario_1" in content:
            msg += content["scenario_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 3", msg)

        msg = ""
        if "scenario_2" in content:
            msg += content["scenario_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 4", msg)

        msg = ""
        if "scenario_3" in content:
            msg += content["scenario_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 5", msg)

        msg = ""
        if "scenario_4" in content:
            msg += content["scenario_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 7", msg)

        msg = ""
        if "scenario_5" in content:
            msg += content["scenario_5"] + "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("Rectangle 8", msg)

        msg = ""
        if "scenario_6" in content:
            msg += content["scenario_6"] + "\n"
        if "detail_6" in content:
            msg += content["detail_6"]
        self._set_text("Rectangle 9", msg)

    def fill_slide_type_92(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "scenario_1" in content:
            self._set_text("ZoneTexte 11", content["scenario_1"])

        if "detail_1" in content:
            self._set_text("TextBox 7", content["detail_1"])

        if "scenario_2" in content:
            self._set_text("Rectangle 11-2", content["scenario_2"])

        if "detail_2" in content:
            self._set_text("TextBox 8", content["detail_2"])

        if "scenario_3" in content:
            self._set_text("Rectangle 11-3", content["scenario_3"])

        if "detail_3" in content:
            self._set_text("TextBox 9", content["detail_3"])

    def fill_slide_type_93(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "scenario_1" in content:
            self._set_text("ZoneTexte 11", content["scenario_1"])

        if "detail_1" in content:
            self._set_text("TextBox 7", content["detail_1"])

        if "scenario_2" in content:
            self._set_text("Rectangle 11-2", content["scenario_2"])

        if "detail_2" in content:
            self._set_text("TextBox 8", content["detail_2"])

        if "scenario_3" in content:
            self._set_text("Rectangle 11-3", content["scenario_3"])

        if "detail_3" in content:
            self._set_text("TextBox 9", content["detail_3"])

        if "scenario_4" in content:
            self._set_text("Rectangle 11-4", content["scenario_4"])

        if "detail_4" in content:
            self._set_text("TextBox 13", content["detail_4"])

    def fill_slide_type_95(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "scenario_1" in content:
            self._set_table_cell("Table 2", 1, 1, content["scenario_1"])

        if "scenario_2" in content:
            self._set_table_cell("Table 2", 1, 2, content["scenario_2"])

        if "characteristic_1" in content:
            self._set_table_cell("Table 2", 2, 1, content["characteristic_1"])

        if "detail_1" in content:
            self._set_table_cell("Table 2", 3, 1, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 2", 3, 2, content["detail_2"])

        if "characteristic_2" in content:
            self._set_table_cell("Table 2", 4, 1, content["characteristic_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 2", 5, 1, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Table 2", 5, 2, content["detail_4"])

        if "characteristic_3" in content:
            self._set_table_cell("Table 2", 6, 1, content["characteristic_3"])

        if "detail_5" in content:
            self._set_table_cell("Table 2", 7, 1, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Table 2", 7, 2, content["detail_6"])

        if "comments" in content:
            self._set_text("Content Placeholder 4", content["comments"])

    def fill_slide_type_96(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "scenario_1" in content:
            self._set_table_cell("Table 2", 1, 1, content["scenario_1"])

        if "scenario_2" in content:
            self._set_table_cell("Table 2", 1, 2, content["scenario_2"])

        if "scenario_3" in content:
            self._set_table_cell("Table 2", 1, 3, content["scenario_3"])

        if "characteristic_1" in content:
            self._set_table_cell("Table 2", 2, 1, content["characteristic_1"])

        if "detail_1" in content:
            self._set_table_cell("Table 2", 3, 1, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 2", 3, 2, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 2", 3, 3, content["detail_3"])

        if "characteristic_2" in content:
            self._set_table_cell("Table 2", 4, 1, content["characteristic_2"])

        if "detail_4" in content:
            self._set_table_cell("Table 2", 5, 1, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Table 2", 5, 2, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Table 2", 5, 3, content["detail_6"])

        if "characteristic_3" in content:
            self._set_table_cell("Table 2", 6, 1, content["characteristic_3"])

        if "detail_7" in content:
            self._set_table_cell("Table 2", 7, 1, content["detail_7"])

        if "detail_8" in content:
            self._set_table_cell("Table 2", 7, 2, content["detail_8"])

        if "detail_9" in content:
            self._set_table_cell("Table 2", 7, 3, content["detail_9"])

        if "comments" in content:
            self._set_text("Content Placeholder 4", content["comments"])

    def fill_slide_type_100(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 2", title)

        if "comments" in content:
            self._set_text("Content Placeholder 13", content["comments"])

        if "in_scope" in content:
            self._set_table_cell("Table 1", 2, 1, content["in_scope"])

        if "in_scope" in content:
            self._set_table_cell("Table 1", 2, 2, content["in_scope"])

    def fill_slide_type_101(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "theme_1" in content:
            self._set_table_cell("Table 5", 2, 1, content["theme_1"])

        if "theme_2" in content:
            self._set_table_cell("Table 5", 3, 1, content["theme_2"])

        if "criteria_1" in content:
            self._set_table_cell("Table 5", 1, 2, content["criteria_1"])

        if "criteria_2" in content:
            self._set_table_cell("Table 5", 1, 3, content["criteria_2"])

        if "detail_1" in content:
            self._set_table_cell("Table 5", 2, 2, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 5", 2, 3, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 5", 3, 2, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Table 5", 3, 3, content["detail_4"])

    def fill_slide_type_102(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "theme_1" in content:
            self._set_table_cell("Table 5", 2, 1, content["theme_1"])

        if "theme_2" in content:
            self._set_table_cell("Table 5", 3, 1, content["theme_2"])

        if "theme_3" in content:
            self._set_table_cell("Table 5", 4, 1, content["theme_3"])

        if "criteria_1" in content:
            self._set_table_cell("Table 5", 1, 2, content["criteria_1"])

        if "criteria_2" in content:
            self._set_table_cell("Table 5", 1, 3, content["criteria_2"])

        if "detail_1" in content:
            self._set_table_cell("Table 5", 2, 2, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 5", 2, 3, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 5", 3, 2, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Table 5", 3, 3, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Table 5", 4, 2, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Table 5", 4, 3, content["detail_6"])

    def fill_slide_type_103(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "theme_1" in content:
            self._set_table_cell("Table 5", 2, 1, content["theme_1"])

        if "theme_2" in content:
            self._set_table_cell("Table 5", 3, 1, content["theme_2"])

        if "theme_3" in content:
            self._set_table_cell("Table 5", 4, 1, content["theme_3"])

        if "theme_4" in content:
            self._set_table_cell("Table 5", 5, 1, content["theme_4"])

        if "criteria_1" in content:
            self._set_table_cell("Table 5", 1, 2, content["criteria_1"])

        if "criteria_2" in content:
            self._set_table_cell("Table 5", 1, 3, content["criteria_2"])

        if "detail_1" in content:
            self._set_table_cell("Table 5", 2, 2, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 5", 2, 3, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 5", 3, 2, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Table 5", 3, 3, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Table 5", 4, 2, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Table 5", 4, 3, content["detail_6"])

        if "detail_7" in content:
            self._set_table_cell("Table 5", 5, 2, content["detail_7"])

        if "detail_8" in content:
            self._set_table_cell("Table 5", 5, 3, content["detail_8"])

    def fill_slide_type_104(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "theme_1" in content:
            self._set_table_cell("Table 5", 2, 1, content["theme_1"])

        if "theme_2" in content:
            self._set_table_cell("Table 5", 3, 1, content["theme_2"])

        if "criteria_1" in content:
            self._set_table_cell("Table 5", 1, 2, content["criteria_1"])

        if "criteria_2" in content:
            self._set_table_cell("Table 5", 1, 3, content["criteria_2"])

        if "criteria_3" in content:
            self._set_table_cell("Table 5", 1, 4, content["criteria_3"])

        if "detail_1" in content:
            self._set_table_cell("Table 5", 2, 2, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 5", 2, 3, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 5", 2, 4, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Table 5", 3, 2, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Table 5", 3, 3, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Table 5", 3, 4, content["detail_6"])

    def fill_slide_type_105(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "theme_1" in content:
            self._set_table_cell("Table 5", 2, 1, content["theme_1"])

        if "theme_2" in content:
            self._set_table_cell("Table 5", 3, 1, content["theme_2"])

        if "theme_3" in content:
            self._set_table_cell("Table 5", 4, 1, content["theme_3"])

        if "criteria_1" in content:
            self._set_table_cell("Table 5", 1, 2, content["criteria_1"])

        if "criteria_2" in content:
            self._set_table_cell("Table 5", 1, 3, content["criteria_2"])

        if "criteria_3" in content:
            self._set_table_cell("Table 5", 1, 4, content["criteria_3"])

        if "detail_1" in content:
            self._set_table_cell("Table 5", 2, 2, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 5", 2, 3, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 5", 2, 4, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Table 5", 3, 2, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Table 5", 3, 3, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Table 5", 3, 4, content["detail_6"])

        if "detail_7" in content:
            self._set_table_cell("Table 5", 4, 2, content["detail_7"])

        if "detail_8" in content:
            self._set_table_cell("Table 5", 4, 3, content["detail_8"])

        if "detail_9" in content:
            self._set_table_cell("Table 5", 4, 4, content["detail_9"])

    def fill_slide_type_106(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "theme_1" in content:
            self._set_table_cell("Table 5", 2, 1, content["theme_1"])

        if "theme_2" in content:
            self._set_table_cell("Table 5", 3, 1, content["theme_2"])

        if "theme_3" in content:
            self._set_table_cell("Table 5", 4, 1, content["theme_3"])

        if "theme_4" in content:
            self._set_table_cell("Table 5", 5, 1, content["theme_4"])

        if "criteria_1" in content:
            self._set_table_cell("Table 5", 1, 2, content["criteria_1"])

        if "criteria_2" in content:
            self._set_table_cell("Table 5", 1, 3, content["criteria_2"])

        if "criteria_3" in content:
            self._set_table_cell("Table 5", 1, 4, content["criteria_3"])

        if "detail_1" in content:
            self._set_table_cell("Table 5", 2, 2, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Table 5", 2, 3, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Table 5", 2, 4, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Table 5", 3, 2, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Table 5", 3, 3, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Table 5", 3, 4, content["detail_6"])

        if "detail_7" in content:
            self._set_table_cell("Table 5", 4, 2, content["detail_7"])

        if "detail_8" in content:
            self._set_table_cell("Table 5", 4, 3, content["detail_8"])

        if "detail_9" in content:
            self._set_table_cell("Table 5", 4, 4, content["detail_9"])

        if "detail_10" in content:
            self._set_table_cell("Table 5", 5, 2, content["detail_10"])

        if "detail_11" in content:
            self._set_table_cell("Table 5", 5, 3, content["detail_11"])

        if "detail_12" in content:
            self._set_table_cell("Table 5", 5, 4, content["detail_12"])

    def fill_slide_type_107(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        msg = ""
        if "idea_1" in content:
            msg += "1. " + content["idea_1"] + "\n"
        if "description_1" in content:
            msg += content["description_1"]
        self._set_table_cell("Tableau 18", 2, 1, msg)

        msg = ""
        if "idea_2" in content:
            msg += "2. " + content["idea_2"] + "\n"
        if "description_2" in content:
            msg += content["description_2"]
        self._set_table_cell("Tableau 18", 3, 1, msg)

        if "pro_1" in content:
            self._set_table_cell("Tableau 18", 2, 2, content["pro_1"])

        if "con_1" in content:
            self._set_table_cell("Tableau 18", 2, 3, content["con_1"])

        if "pro_2" in content:
            self._set_table_cell("Tableau 18", 3, 2, content["pro_2"])

        if "con_2" in content:
            self._set_table_cell("Tableau 18", 3, 3, content["con_2"])

    def fill_slide_type_108(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        msg = ""
        if "idea_1" in content:
            msg += "1. " + content["idea_1"] + "\n"
        if "description_1" in content:
            msg += content["description_1"]
        self._set_table_cell("Tableau 18", 2, 1, msg)

        msg = ""
        if "idea_2" in content:
            msg += "2. " + content["idea_2"] + "\n"
        if "description_2" in content:
            msg += content["description_2"]
        self._set_table_cell("Tableau 18", 3, 1, msg)

        msg = ""
        if "idea_3" in content:
            msg += "3. " + content["idea_3"] + "\n"
        if "description_3" in content:
            msg += content["description_3"]
        self._set_table_cell("Tableau 18", 4, 1, msg)

        if "pro_1" in content:
            self._set_table_cell("Tableau 18", 2, 2, content["pro_1"])

        if "con_1" in content:
            self._set_table_cell("Tableau 18", 2, 3, content["con_1"])

        if "pro_2" in content:
            self._set_table_cell("Tableau 18", 3, 2, content["pro_2"])

        if "con_2" in content:
            self._set_table_cell("Tableau 18", 3, 3, content["con_2"])

        if "pro_3" in content:
            self._set_table_cell("Tableau 18", 4, 2, content["pro_3"])

        if "con_3" in content:
            self._set_table_cell("Tableau 18", 4, 3, content["con_3"])

    def fill_slide_type_109(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        msg = ""
        if "idea_1" in content:
            msg += "1. " + content["idea_1"] + "\n"
        if "description_1" in content:
            msg += content["description_1"]
        self._set_table_cell("Tableau 18", 2, 1, msg)

        msg = ""
        if "idea_2" in content:
            msg += "2. " + content["idea_2"] + "\n"
        if "description_2" in content:
            msg += content["description_2"]
        self._set_table_cell("Tableau 18", 3, 1, msg)

        msg = ""
        if "idea_3" in content:
            msg += "3. " + content["idea_3"] + "\n"
        if "description_3" in content:
            msg += content["description_3"]
        self._set_table_cell("Tableau 18", 4, 1, msg)

        msg = ""
        if "idea_4" in content:
            msg += "4. " + content["idea_4"] + "\n"
        if "description_4" in content:
            msg += content["description_4"]
        self._set_table_cell("Tableau 18", 5, 1, msg)

        if "pro_1" in content:
            self._set_table_cell("Tableau 18", 2, 2, content["pro_1"])

        if "con_1" in content:
            self._set_table_cell("Tableau 18", 2, 3, content["con_1"])

        if "pro_2" in content:
            self._set_table_cell("Tableau 18", 3, 2, content["pro_2"])

        if "con_2" in content:
            self._set_table_cell("Tableau 18", 3, 3, content["con_2"])

        if "pro_3" in content:
            self._set_table_cell("Tableau 18", 4, 2, content["pro_3"])

        if "con_3" in content:
            self._set_table_cell("Tableau 18", 4, 3, content["con_3"])

        if "pro_4" in content:
            self._set_table_cell("Tableau 18", 5, 2, content["pro_4"])

        if "con_4" in content:
            self._set_table_cell("Tableau 18", 5, 3, content["con_4"])

    def fill_slide_type_110(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "theme_1" in content:
            self._set_table_cell("Tableau 5", 1, 1, content["theme_1"])

        if "theme_2" in content:
            self._set_table_cell("Tableau 5", 1, 2, content["theme_2"])

        if "detail_1" in content:
            self._set_table_cell("Tableau 5", 2, 1, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Tableau 5", 2, 2, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Tableau 5", 3, 1, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Tableau 5", 3, 2, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Tableau 5", 4, 1, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Tableau 5", 4, 2, content["detail_6"])

    def fill_slide_type_112(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "kpi_1" in content and "kpi_1_unit" in content:
            self._set_table_cell(
                "Table 7", 1, 3, content["kpi_1"] + "\n" + content["kpi_1_unit"])

        if "kpi_2" in content and "kpi_2_unit" in content:
            self._set_table_cell(
                "Table 7", 1, 4, content["kpi_2"] + "\n" + content["kpi_2_unit"])

        if "scenario_1" in content:
            self._set_table_cell("Table 7", 2, 1, content["scenario_1"])

        if "scenario_2" in content:
            self._set_table_cell("Table 7", 3, 2, content["scenario_2"])

        if "scenario_3" in content:
            self._set_table_cell("Table 7", 4, 2, content["scenario_3"])

    def fill_slide_type_113(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "kpi_1" in content and "kpi_1_unit" in content:
            self._set_table_cell(
                "Table 7", 1, 3, content["kpi_1"] + "\n" + content["kpi_1_unit"])

        if "kpi_2" in content and "kpi_2_unit" in content:
            self._set_table_cell(
                "Table 7", 1, 4, content["kpi_2"] + "\n" + content["kpi_2_unit"])

        if "scenario_1" in content:
            self._set_table_cell("Table 7", 2, 1, content["scenario_1"])

        if "scenario_2" in content:
            self._set_table_cell("Table 7", 3, 2, content["scenario_2"])

        if "scenario_3" in content:
            self._set_table_cell("Table 7", 4, 2, content["scenario_3"])

        if "scenario_4" in content:
            self._set_table_cell("Table 7", 5, 2, content["scenario_4"])

    def fill_slide_type_115(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "phase_1" in content:
            self._set_text("Rectangle 9", content["phase_1"])

        if "phase_2" in content:
            self._set_text("Rectangle 23", content["phase_2"])

        if "phase_3" in content:
            self._set_text("Rectangle 25", content["phase_3"])

        if "phase_4" in content:
            self._set_text("Rectangle 27", content["phase_4"])

    def fill_slide_type_116(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 8", title)

        msg = ""
        if "phase_1" in content:
            msg += content["phase_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 19", msg)

        msg = ""
        if "phase_2" in content:
            msg += content["phase_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 10", msg)

        msg = ""
        if "phase_3" in content:
            msg += content["phase_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 11", msg)

        msg = ""
        if "phase_4" in content:
            msg += content["phase_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 68", msg)

    def fill_slide_type_117(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 8", title)

        msg = ""
        if "phase_1" in content:
            msg += content["phase_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 20", msg)

        msg = ""
        if "phase_2" in content:
            msg += content["phase_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 10", msg)

        msg = ""
        if "phase_3" in content:
            msg += content["phase_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 11", msg)

        msg = ""
        if "phase_4" in content:
            msg += content["phase_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 68", msg)

        msg = ""
        if "phase_5" in content:
            msg += content["phase_5"] + "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("Rectangle 27", msg)

    def fill_slide_type_118(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 8", title)

        msg = ""
        if "phase_1" in content:
            msg += content["phase_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("Rectangle 26", msg)

        msg = ""
        if "phase_2" in content:
            msg += content["phase_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("Rectangle 10", msg)

        msg = ""
        if "phase_3" in content:
            msg += content["phase_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("Rectangle 11", msg)

        msg = ""
        if "phase_4" in content:
            msg += content["phase_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("Rectangle 68", msg)

        msg = ""
        if "phase_5" in content:
            msg += content["phase_5"] + "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("Rectangle 21", msg)

        msg = ""
        if "phase_6" in content:
            msg += content["phase_6"] + "\n"
        if "detail_6" in content:
            msg += content["detail_6"]
        self._set_text("Rectangle 23", msg)

    def fill_slide_type_119(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 9", title)

        if "step_1" in content:
            self._set_text("Rectangle 32", content["step_1"])

        if "deliverable_1" in content:
            self._set_text("Rectangle 24", content["deliverable_1"])

        if "step_2" in content:
            self._set_text("Rectangle 35", content["step_2"])

        if "deliverable_2" in content:
            self._set_text("Rectangle 34", content["deliverable_2"])

        if "step_3" in content:
            self._set_text("Rectangle 36", content["step_3"])

        if "deliverable_3" in content:
            self._set_text("Rectangle 44", content["deliverable_3"])

    def fill_slide_type_120(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "step_1" in content:
            self._set_text("Rectangle 42", content["step_1"])

        if "deliverable_1" in content:
            self._set_text("Rectangle 41", content["deliverable_1"])

        if "step_2" in content:
            self._set_text("Rectangle 43", content["step_2"])

        if "deliverable_2" in content:
            self._set_text("Rectangle 46", content["deliverable_2"])

        if "step_3" in content:
            self._set_text("Rectangle 45", content["step_3"])

        if "deliverable_3" in content:
            self._set_text("Rectangle 48", content["deliverable_3"])

        if "step_4" in content:
            self._set_text("Rectangle 55", content["step_4"])

        if "deliverable_4" in content:
            self._set_text("Rectangle 56", content["deliverable_4"])

    def fill_slide_type_121(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 4", title)

        if "step_1" in content:
            self._set_text("TextBox 9", "Step 1\n" + content["step_1"])

        if "step_2" in content:
            self._set_text("TextBox 12", "Step 2\n" + content["step_2"])

        if "step_3" in content:
            self._set_text("TextBox 10", "Step 3\n" + content["step_3"])

        if "step_4" in content:
            self._set_text("TextBox 11", "Step 4\n" + content["step_4"])

    def fill_slide_type_122(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "step_1" in content:
            self._set_text("Rectangle 51", content["step_1"])

        if "step_2" in content:
            self._set_text("Rectangle 56", content["step_2"])

        if "step_3" in content:
            self._set_text("Rectangle 57", content["step_3"])

    def fill_slide_type_123(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "step_1" in content:
            self._set_text("Rectangle 51", content["step_1"])

        if "step_2" in content:
            self._set_text("Rectangle 56", content["step_2"])

        if "step_3" in content:
            self._set_text("Rectangle 57", content["step_3"])

    def fill_slide_type_124(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "step_1" in content:
            self._set_text("Rectangle 51", content["step_1"])

        if "step_2" in content:
            self._set_text("Rectangle 56", content["step_2"])

        if "step_3" in content:
            self._set_text("Rectangle 57", content["step_3"])

        if "step_4" in content:
            self._set_text("Rectangle 58", content["step_4"])

    def fill_slide_type_125(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "step_1" in content:
            self._set_text("Rectangle 51", content["step_1"])

        if "step_2" in content:
            self._set_text("Rectangle 56", content["step_2"])

        if "step_3" in content:
            self._set_text("Rectangle 57", content["step_3"])

        if "step_4" in content:
            self._set_text("Rectangle 58", content["step_4"])

    def fill_slide_type_126(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "step_1" in content:
            self._set_text("Rectangle 51", content["step_1"])

        if "step_2" in content:
            self._set_text("Rectangle 56", content["step_2"])

        if "step_3" in content:
            self._set_text("Rectangle 57", content["step_3"])

        if "step_4" in content:
            self._set_text("Rectangle 58", content["step_4"])

        if "step_5" in content:
            self._set_text("Rectangle 59", content["step_5"])

    def fill_slide_type_127(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 4", title)

        msg = ""
        if "step_1" in content:
            msg += content["step_1"] + "\n"
        if "deliverable_1" in content:
            msg += content["deliverable_1"]
        self._set_text("TextBox 66", msg)

        msg = ""
        if "step_2" in content:
            msg += content["step_2"] + "\n"
        if "deliverable_2" in content:
            msg += content["deliverable_2"]
        self._set_text("TextBox 68", msg)

        msg = ""
        if "step_3" in content:
            msg += content["step_3"] + "\n"
        if "deliverable_3" in content:
            msg += content["deliverable_3"]
        self._set_text("TextBox 65", msg)

        msg = ""
        if "step_4" in content:
            msg += content["step_4"] + "\n"
        if "deliverable_4" in content:
            msg += content["deliverable_4"]
        self._set_text("TextBox 67", msg)

        msg = ""
        if "step_5" in content:
            msg += content["step_5"] + "\n"
        if "deliverable_5" in content:
            msg += content["deliverable_5"]
        self._set_text("TextBox 63", msg)

    def fill_slide_type_128(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 71", title)

        if "step_1" in content:
            self._set_group_text("Groupe 21", "TextBox 65", content["step_1"])

        if "description_1" in content:
            self._set_text("TextBox 59-1", content["description_1"])

        if "step_2" in content:
            self._set_group_text("Groupe 21", "TextBox 68", content["step_2"])

        if "description_2" in content:
            self._set_text("TextBox 59-2", content["description_2"])

        if "step_3" in content:
            self._set_group_text("Groupe 21", "TextBox 66", content["step_3"])

        if "description_3" in content:
            self._set_text("TextBox 59-3", content["description_3"])

        if "step_4" in content:
            self._set_group_text("Groupe 21", "TextBox 69", content["step_4"])

        if "description_4" in content:
            self._set_text("TextBox 59-4", content["description_4"])

        if "step_5" in content:
            self._set_group_text("Groupe 21", "TextBox 67", content["step_5"])

        if "description_5" in content:
            self._set_text("TextBox 59-5", content["description_5"])

        if "step_6" in content:
            self._set_group_text("Groupe 21", "TextBox 70", content["step_6"])

        if "description_6" in content:
            self._set_text("TextBox 59-6", content["description_6"])

    def fill_slide_type_129(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 71", title)

        if "step_1" in content:
            self._set_group_text("Groupe 21", "TextBox 65", content["step_1"])

        if "description_1" in content:
            self._set_text("TextBox 59-1", content["description_1"])

        if "step_2" in content:
            self._set_group_text("Groupe 21", "TextBox 68", content["step_2"])

        if "description_2" in content:
            self._set_text("TextBox 59-2", content["description_2"])

        if "step_3" in content:
            self._set_group_text("Groupe 21", "TextBox 66", content["step_3"])

        if "description_3" in content:
            self._set_text("TextBox 59-3", content["description_3"])

    def fill_slide_type_130(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 71", title)

        if "step_1" in content:
            self._set_group_text("Groupe 21", "TextBox 65", content["step_1"])

        if "description_1" in content:
            self._set_text("TextBox 59-1", content["description_1"])

        if "step_2" in content:
            self._set_group_text("Groupe 21", "TextBox 68", content["step_2"])

        if "description_2" in content:
            self._set_text("TextBox 59-2", content["description_2"])

        if "step_3" in content:
            self._set_group_text("Groupe 21", "TextBox 66", content["step_3"])

        if "description_3" in content:
            self._set_text("TextBox 59-3", content["description_3"])

        if "step_4" in content:
            self._set_group_text("Groupe 21", "TextBox 69", content["step_4"])

        if "description_4" in content:
            self._set_text("TextBox 59-4", content["description_4"])

    def fill_slide_type_131(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 71", title)

        if "step_1" in content:
            self._set_text("TextBox 65", content["step_1"])

        if "description_1" in content:
            self._set_text("TextBox 59-1", content["description_1"])

        if "step_2" in content:
            self._set_text("TextBox 68", content["step_2"])

        if "description_2" in content:
            self._set_text("TextBox 59-2", content["description_2"])

        if "step_3" in content:
            self._set_text("TextBox 66", content["step_3"])

        if "description_3" in content:
            self._set_text("TextBox 59-3", content["description_3"])

        if "step_4" in content:
            self._set_text("TextBox 69", content["step_4"])

        if "description_4" in content:
            self._set_text("TextBox 59-4", content["description_4"])

        if "step_5" in content:
            self._set_text("TextBox 67", content["step_5"])

        if "description_5" in content:
            self._set_text("TextBox 59-5", content["description_5"])

    def fill_slide_type_135(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "step_1" in content:
            self._set_text("Flèche : pentagone 3", content["step_1"])

        msg = ""
        if "deliverable_1" in content:
            msg += content["deliverable_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("ZoneTexte 24", msg)

        msg = ""
        if "deliverable_2" in content:
            msg += content["deliverable_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("ZoneTexte 28", msg)

        if "step_2" in content:
            self._set_text("Flèche : chevron 4", content["step_2"])

        msg = ""
        if "deliverable_3" in content:
            msg += content["deliverable_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("ZoneTexte 29", msg)

        msg = ""
        if "deliverable_4" in content:
            msg += content["deliverable_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("ZoneTexte 30", msg)

        if "step_3" in content:
            self._set_text("Flèche : chevron 5", content["step_3"])

        msg = ""
        if "deliverable_5" in content:
            msg += content["deliverable_5"] + "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("ZoneTexte 33", msg)

        msg = ""
        if "deliverable_6" in content:
            msg += content["deliverable_6"] + "\n"
        if "detail_6" in content:
            msg += content["detail_6"]
        self._set_text("ZoneTexte 34", msg)

    def fill_slide_type_136(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "step_1" in content:
            self._set_text("Flèche : pentagone 3", content["step_1"])

        msg = ""
        if "deliverable_1" in content:
            msg += content["deliverable_1"] + "\n"
        if "detail_1" in content:
            msg += content["detail_1"]
        self._set_text("ZoneTexte 24", msg)

        msg = ""
        if "deliverable_2" in content:
            msg += content["deliverable_2"] + "\n"
        if "detail_2" in content:
            msg += content["detail_2"]
        self._set_text("ZoneTexte 28", msg)

        if "step_2" in content:
            self._set_text("Flèche : chevron 4", content["step_2"])

        msg = ""
        if "deliverable_3" in content:
            msg += content["deliverable_3"] + "\n"
        if "detail_3" in content:
            msg += content["detail_3"]
        self._set_text("ZoneTexte 80", msg)

        msg = ""
        if "deliverable_4" in content:
            msg += content["deliverable_4"] + "\n"
        if "detail_4" in content:
            msg += content["detail_4"]
        self._set_text("ZoneTexte 81", msg)

        if "step_3" in content:
            self._set_text("Flèche : chevron 5", content["step_3"])

        msg = ""
        if "deliverable_5" in content:
            msg += content["deliverable_5"] + "\n"
        if "detail_5" in content:
            msg += content["detail_5"]
        self._set_text("ZoneTexte 82", msg)

        msg = ""
        if "deliverable_6" in content:
            msg += content["deliverable_6"] + "\n"
        if "detail_6" in content:
            msg += content["detail_6"]
        self._set_text("ZoneTexte 83", msg)

        if "step_4" in content:
            self._set_text("Flèche : chevron 65", content["step_4"])

        msg = ""
        if "deliverable_7" in content:
            msg += content["deliverable_7"] + "\n"
        if "detail_7" in content:
            msg += content["detail_7"]
        self._set_text("ZoneTexte 84", msg)

        msg = ""
        if "deliverable_8" in content:
            msg += content["deliverable_8"] + "\n"
        if "detail_8" in content:
            msg += content["detail_8"]
        self._set_text("ZoneTexte 85", msg)

    def fill_slide_type_137(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "phase_1" in content:
            self._set_text("Pentagone 17", content["phase_1"])

        if "phase_2" in content:
            self._set_text("Chevron 10", content["phase_2"])

        if "phase_3" in content:
            self._set_text("Chevron 24", content["phase_3"])

        if "characteristic_1" in content:
            self._set_table_cell("Content Placeholder 3",
                                 1, 1, content["characteristic_1"])

        if "characteristic_2" in content:
            self._set_table_cell("Content Placeholder 3",
                                 3, 1, content["characteristic_2"])

        if "characteristic_3" in content:
            self._set_table_cell("Content Placeholder 3",
                                 5, 1, content["characteristic_3"])

        if "description_1" in content:
            self._set_table_cell("Content Placeholder 3",
                                 2, 1, content["description_1"])

        if "description_2" in content:
            self._set_table_cell("Content Placeholder 3",
                                 2, 2, content["description_2"])

        if "description_3" in content:
            self._set_table_cell("Content Placeholder 3",
                                 2, 3, content["description_3"])

        if "detail_1" in content:
            self._set_table_cell("Content Placeholder 3",
                                 4, 1, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Content Placeholder 3",
                                 4, 2, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Content Placeholder 3",
                                 4, 3, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Content Placeholder 3",
                                 6, 1, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Content Placeholder 3",
                                 6, 2, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Content Placeholder 3",
                                 6, 3, content["detail_6"])

    def fill_slide_type_138(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 1", title)

        if "phase_1" in content:
            self._set_text("Pentagone 17", content["phase_1"])

        if "phase_2" in content:
            self._set_text("Chevron 10", content["phase_2"])

        if "phase_3" in content:
            self._set_text("Chevron 24", content["phase_3"])

        if "phase_4" in content:
            self._set_text("Chevron 24-2", content["phase_4"])

        if "characteristic_1" in content:
            self._set_table_cell("Content Placeholder 3",
                                 1, 1, content["characteristic_1"])

        if "characteristic_2" in content:
            self._set_table_cell("Content Placeholder 3",
                                 3, 1, content["characteristic_2"])

        if "characteristic_3" in content:
            self._set_table_cell("Content Placeholder 3",
                                 5, 1, content["characteristic_3"])

        if "description_1" in content:
            self._set_table_cell("Content Placeholder 3",
                                 2, 1, content["description_1"])

        if "description_2" in content:
            self._set_table_cell("Content Placeholder 3",
                                 2, 2, content["description_2"])

        if "description_3" in content:
            self._set_table_cell("Content Placeholder 3",
                                 2, 3, content["description_3"])

        if "description_4" in content:
            self._set_table_cell("Content Placeholder 3",
                                 2, 4, content["description_4"])

        if "detail_1" in content:
            self._set_table_cell("Content Placeholder 3",
                                 4, 1, content["detail_1"])

        if "detail_2" in content:
            self._set_table_cell("Content Placeholder 3",
                                 4, 2, content["detail_2"])

        if "detail_3" in content:
            self._set_table_cell("Content Placeholder 3",
                                 4, 3, content["detail_3"])

        if "detail_4" in content:
            self._set_table_cell("Content Placeholder 3",
                                 4, 4, content["detail_4"])

        if "detail_5" in content:
            self._set_table_cell("Content Placeholder 3",
                                 6, 1, content["detail_5"])

        if "detail_6" in content:
            self._set_table_cell("Content Placeholder 3",
                                 6, 2, content["detail_6"])

        if "detail_7" in content:
            self._set_table_cell("Content Placeholder 3",
                                 6, 3, content["detail_7"])

        if "detail_8" in content:
            self._set_table_cell("Content Placeholder 3",
                                 6, 4, content["detail_8"])

    def fill_slide_type_139(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 3", title)

        if "phase_1" in content:
            self._set_text("Rectangle 89", content["phase_1"])

        if "phase_2" in content:
            self._set_text("Rectangle 68", content["phase_2"])

        if "phase_3" in content:
            self._set_text("Rectangle 61", content["phase_3"])

        if "characteristic_1" in content:
            self._set_text("ZoneTexte 32", content["characteristic_1"])

        if "characteristic_2" in content:
            self._set_text("ZoneTexte 33", content["characteristic_2"])

        if "characteristic_3" in content:
            self._set_text("ZoneTexte 36", content["characteristic_3"])

        if "detail_1" in content:
            self._set_text("ZoneTexte 43", content["detail_1"])

        if "detail_2" in content:
            self._set_text("ZoneTexte 44", content["detail_2"])

        if "detail_3" in content:
            self._set_text("ZoneTexte 45", content["detail_3"])

        if "detail_4" in content:
            self._set_text("ZoneTexte 40", content["detail_4"])

        if "detail_5" in content:
            self._set_text("ZoneTexte 41", content["detail_5"])

        if "detail_6" in content:
            self._set_text("ZoneTexte 42", content["detail_6"])

        if "detail_7" in content:
            self._set_text("ZoneTexte 49", content["detail_7"])

        if "detail_8" in content:
            self._set_text("ZoneTexte 47", content["detail_8"])

        if "detail_9" in content:
            self._set_text("ZoneTexte 46", content["detail_9"])

    def fill_slide_type_143(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "step_number" in content:
            self._set_text("Google Shape;2350;p252",
                           "Step #" + content["step_number"])

        if "step_title" in content:
            self._set_text("Google Shape;2352;p252", content["step_title"])

        if "step_description" in content:
            self._set_text("Content Placeholder 5",
                           content["step_description"])

        if "ksf_1_title" in content and "ksf_1_description" in content:
            self._set_text("Google Shape;2359;p252",
                           content["ksf_1_title"] + "\n" + content["ksf_1_description"])

        if "ksf_2_title" in content and "ksf_2_description" in content:
            self._set_text("Google Shape;2359;p252-2",
                           content["ksf_2_title"] + "\n" + content["ksf_2_description"])

        if "ksf_3_title" in content and "ksf_3_description" in content:
            self._set_text("Google Shape;2359;p252-3",
                           content["ksf_3_title"] + "\n" + content["ksf_3_description"])

        if "ksf_4_title" in content and "ksf_4_description" in content:
            self._set_text("Google Shape;2359;p252-4",
                           content["ksf_4_title"] + "\n" + content["ksf_4_description"])

    def fill_slide_type_144(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "step_1" in content:
            self._set_text("Google Shape;2260;p246-1",
                           "1\n" + content["step_1"])

        if "step_2" in content:
            self._set_text("Google Shape;2260;p246-2",
                           "2\n" + content["step_2"])

        if "step_3" in content:
            self._set_text("Google Shape;2260;p246-3",
                           "3\n" + content["step_3"])

        if "step_subtitle" in content and "step_description" in content:
            self._set_text("Google Shape;2262;p246",
                           content["step_subtitle"] + "\n\n" + content["step_description"])

    def fill_slide_type_145(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "step_1" in content:
            self._set_text("Google Shape;2260;p246-1",
                           "1\n" + content["step_1"])

        if "step_2" in content:
            self._set_text("Google Shape;2260;p246-2",
                           "2\n" + content["step_2"])

        if "step_3" in content:
            self._set_text("Google Shape;2260;p246-3",
                           "3\n" + content["step_3"])

        if "step_subtitle" in content and "step_description" in content:
            self._set_text("Google Shape;2262;p246",
                           content["step_subtitle"] + "\n\n" + content["step_description"])

    def fill_slide_type_146(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "step_1" in content:
            self._set_text("Google Shape;2260;p246-1",
                           "1\n" + content["step_1"])

        if "step_2" in content:
            self._set_text("Google Shape;2260;p246-2",
                           "2\n" + content["step_2"])

        if "step_3" in content:
            self._set_text("Google Shape;2260;p246-3",
                           "3\n" + content["step_3"])

        if "step_subtitle" in content and "step_description" in content:
            self._set_text("Google Shape;2262;p246",
                           content["step_subtitle"] + "\n\n" + content["step_description"])

    def fill_slide_type_147(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "description_1" in content:
            self._set_text("Rectangle 13", content["description_1"])

        if "description_2" in content:
            self._set_text("Rectangle 68", content["description_2"])

        if "description_3" in content:
            self._set_text("Rectangle 72", content["description_3"])

        if "description_4" in content:
            self._set_text("Rectangle 84", content["description_4"])

    def fill_slide_type_148(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "description_1" in content:
            self._set_text("Rectangle 13", content["description_1"])

        if "description_2" in content:
            self._set_text("Rectangle 68", content["description_2"])

        if "description_3" in content:
            self._set_text("Rectangle 72", content["description_3"])

        if "description_4" in content:
            self._set_text("Rectangle 76", content["description_4"])

        if "description_5" in content:
            self._set_text("Rectangle 84", content["description_5"])

    def fill_slide_type_149(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Title 12", title)

        if "description_1" in content:
            self._set_text("Rectangle 13", content["description_1"])

        if "description_2" in content:
            self._set_text("Rectangle 68", content["description_2"])

        if "description_3" in content:
            self._set_text("Rectangle 72", content["description_3"])

        if "description_4" in content:
            self._set_text("Rectangle 76", content["description_4"])

        if "description_5" in content:
            self._set_text("Rectangle 80", content["description_5"])

        if "description_6" in content:
            self._set_text("Rectangle 84", content["description_6"])

    def fill_slide_type_150(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "week_1_content" in content:
            self._set_text("Rectangle 9", content["week_1_content"])

        if "week_2_content" in content:
            self._set_text("Rectangle 8", content["week_2_content"])

        if "week_3_content" in content:
            self._set_text("Rectangle 6", content["week_3_content"])

        if "final_day_content" in content:
            self._set_text("Rectangle 5", content["final_day_content"])

    def fill_slide_type_151(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "week_1_content" in content:
            self._set_text("Rectangle 10", content["week_1_content"])

        if "week_2_content" in content:
            self._set_text("Rectangle 9", content["week_2_content"])

        if "week_3_content" in content:
            self._set_text("Rectangle 8", content["week_3_content"])

        if "week_4_content" in content:
            self._set_text("Rectangle 6", content["week_4_content"])

        if "final_day_content" in content:
            self._set_text("Rectangle 5", content["final_day_content"])

    def fill_slide_type_152(self, index: int,  title: str, content: Dict[str, str]):
        self.slide = self.presentation.Slides(index)
        self._build_shape_cache()

        self._set_text("Titre 1", title)

        if "success_key_1" in content and "description_1" in content:
            self._set_text(
                "Rectangle 24", content["success_key_1"] + "\n" + content["description_1"])

        if "success_key_2" in content and "description_2" in content:
            self._set_text(
                "Rectangle 25", content["success_key_1"] + "\n" + content["description_2"])

        if "success_key_3" in content and "description_3" in content:
            self._set_text(
                "Rectangle 26", content["success_key_3"] + "\n" + content["description_3"])

        if "success_key_4" in content and "description_4" in content:
            self._set_text(
                "Rectangle 27", content["success_key_4"] + "\n" + content["description_4"])

    # def fill_slide_type_154(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #154.

    #     Modifiable elements (12):
    #     - ZoneTexte 47
    #     - ZoneTexte 48
    #     - Titre 1 [TITLE]
    #     - Rectangle 3
    #     - Rectangle 4
    #     - Rectangle 5
    #     - Rectangle 6
    #     - ZoneTexte 44
    #     - ZoneTexte 45
    #     - ZoneTexte 51
    #     - ZoneTexte 52
    #     - Rectangle 28

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(154)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill ZoneTexte 47
    #     if "zonetexte_47" in data:
    #         self._set_text("ZoneTexte 47", data["zonetexte_47"])

    #     # Fill ZoneTexte 48
    #     if "zonetexte_48" in data:
    #         self._set_text("ZoneTexte 48", data["zonetexte_48"])

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill Rectangle 3
    #     if "rectangle_3" in data:
    #         self._set_text("Rectangle 3", data["rectangle_3"])

    #     # Fill Rectangle 4
    #     if "rectangle_4" in data:
    #         self._set_text("Rectangle 4", data["rectangle_4"])

    #     # Fill Rectangle 5
    #     if "rectangle_5" in data:
    #         self._set_text("Rectangle 5", data["rectangle_5"])

    #     # Fill Rectangle 6
    #     if "rectangle_6" in data:
    #         self._set_text("Rectangle 6", data["rectangle_6"])

    #     # Fill ZoneTexte 44
    #     if "zonetexte_44" in data:
    #         self._set_text("ZoneTexte 44", data["zonetexte_44"])

    #     # Fill ZoneTexte 45
    #     if "zonetexte_45" in data:
    #         self._set_text("ZoneTexte 45", data["zonetexte_45"])

    #     # Fill ZoneTexte 51
    #     if "zonetexte_51" in data:
    #         self._set_text("ZoneTexte 51", data["zonetexte_51"])

    #     # Fill ZoneTexte 52
    #     if "zonetexte_52" in data:
    #         self._set_text("ZoneTexte 52", data["zonetexte_52"])

    #     # Fill Rectangle 28
    #     if "rectangle_28" in data:
    #         self._set_text("Rectangle 28", data["rectangle_28"])

    # def fill_slide_type_155(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #155.

    #     Modifiable elements (27):
    #     - Titre 3 [TITLE]
    #     - ZoneTexte 13
    #     - ZoneTexte 14
    #     - Tableau 15
    #     - ZoneTexte 17
    #     - ZoneTexte 18
    #     - ZoneTexte 19
    #     - ZoneTexte 1
    #     - ZoneTexte 12
    #     - ZoneTexte 16
    #     - ZoneTexte 20
    #     - ZoneTexte 71
    #     - Ellipse 52
    #     - Ellipse 53
    #     - Ellipse 54
    #     - Ellipse 55
    #     - Ellipse 58
    #     - Ellipse 64
    #     - Ellipse 72
    #     - Ellipse 73
    #     - Ellipse 74
    #     - Ellipse 75
    #     - Ellipse 76
    #     - Ellipse 77
    #     - Ellipse 78
    #     - Ellipse 79
    #     - Table 3

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(155)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 3
    #     if "title" in data:
    #         self._set_text("Titre 3", data["title"])

    #     # Fill ZoneTexte 13
    #     if "zonetexte_13" in data:
    #         self._set_text("ZoneTexte 13", data["zonetexte_13"])

    #     # Fill ZoneTexte 14
    #     if "zonetexte_14" in data:
    #         self._set_text("ZoneTexte 14", data["zonetexte_14"])

    #     # Fill Tableau 15
    #     if "tableau_15" in data:
    #         self._set_text("Tableau 15", data["tableau_15"])

    #     # Fill ZoneTexte 17
    #     if "zonetexte_17" in data:
    #         self._set_text("ZoneTexte 17", data["zonetexte_17"])

    #     # Fill ZoneTexte 18
    #     if "zonetexte_18" in data:
    #         self._set_text("ZoneTexte 18", data["zonetexte_18"])

    #     # Fill ZoneTexte 19
    #     if "zonetexte_19" in data:
    #         self._set_text("ZoneTexte 19", data["zonetexte_19"])

    #     # Fill ZoneTexte 1
    #     if "zonetexte_1" in data:
    #         self._set_text("ZoneTexte 1", data["zonetexte_1"])

    #     # Fill ZoneTexte 12
    #     if "zonetexte_12" in data:
    #         self._set_text("ZoneTexte 12", data["zonetexte_12"])

    #     # Fill ZoneTexte 16
    #     if "zonetexte_16" in data:
    #         self._set_text("ZoneTexte 16", data["zonetexte_16"])

    #     # Fill ZoneTexte 20
    #     if "zonetexte_20" in data:
    #         self._set_text("ZoneTexte 20", data["zonetexte_20"])

    #     # Fill ZoneTexte 71
    #     if "zonetexte_71" in data:
    #         self._set_text("ZoneTexte 71", data["zonetexte_71"])

    #     # Fill Ellipse 52
    #     if "ellipse_52" in data:
    #         self._set_text("Ellipse 52", data["ellipse_52"])

    #     # Fill Ellipse 53
    #     if "ellipse_53" in data:
    #         self._set_text("Ellipse 53", data["ellipse_53"])

    #     # Fill Ellipse 54
    #     if "ellipse_54" in data:
    #         self._set_text("Ellipse 54", data["ellipse_54"])

    #     # Fill Ellipse 55
    #     if "ellipse_55" in data:
    #         self._set_text("Ellipse 55", data["ellipse_55"])

    #     # Fill Ellipse 58
    #     if "ellipse_58" in data:
    #         self._set_text("Ellipse 58", data["ellipse_58"])

    #     # Fill Ellipse 64
    #     if "ellipse_64" in data:
    #         self._set_text("Ellipse 64", data["ellipse_64"])

    #     # Fill Ellipse 72
    #     if "ellipse_72" in data:
    #         self._set_text("Ellipse 72", data["ellipse_72"])

    #     # Fill Ellipse 73
    #     if "ellipse_73" in data:
    #         self._set_text("Ellipse 73", data["ellipse_73"])

    #     # Fill Ellipse 74
    #     if "ellipse_74" in data:
    #         self._set_text("Ellipse 74", data["ellipse_74"])

    #     # Fill Ellipse 75
    #     if "ellipse_75" in data:
    #         self._set_text("Ellipse 75", data["ellipse_75"])

    #     # Fill Ellipse 76
    #     if "ellipse_76" in data:
    #         self._set_text("Ellipse 76", data["ellipse_76"])

    #     # Fill Ellipse 77
    #     if "ellipse_77" in data:
    #         self._set_text("Ellipse 77", data["ellipse_77"])

    #     # Fill Ellipse 78
    #     if "ellipse_78" in data:
    #         self._set_text("Ellipse 78", data["ellipse_78"])

    #     # Fill Ellipse 79
    #     if "ellipse_79" in data:
    #         self._set_text("Ellipse 79", data["ellipse_79"])

    #     # Fill Table 3
    #     if "table_3" in data:
    #         self._set_text("Table 3", data["table_3"])

    # def fill_slide_type_156(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #156.

    #     Modifiable elements (21):
    #     - ZoneTexte 32
    #     - Titre 2 [TITLE]
    #     - ZoneTexte 13
    #     - ZoneTexte 14
    #     - Tableau 15
    #     - ZoneTexte 1
    #     - ZoneTexte 12
    #     - ZoneTexte 16
    #     - ZoneTexte 20
    #     - ZoneTexte 5
    #     - ZoneTexte 26
    #     - ZoneTexte 28
    #     - Ellipse 72
    #     - Ellipse 74
    #     - Ellipse 75
    #     - ZoneTexte 23
    #     - ZoneTexte 76
    #     - ZoneTexte 78
    #     - ZoneTexte 29
    #     - ZoneTexte 30
    #     - ZoneTexte 31

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(156)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill ZoneTexte 32
    #     if "zonetexte_32" in data:
    #         self._set_text("ZoneTexte 32", data["zonetexte_32"])

    #     # Fill Titre 2
    #     if "title" in data:
    #         self._set_text("Titre 2", data["title"])

    #     # Fill ZoneTexte 13
    #     if "zonetexte_13" in data:
    #         self._set_text("ZoneTexte 13", data["zonetexte_13"])

    #     # Fill ZoneTexte 14
    #     if "zonetexte_14" in data:
    #         self._set_text("ZoneTexte 14", data["zonetexte_14"])

    #     # Fill Tableau 15
    #     if "tableau_15" in data:
    #         self._set_text("Tableau 15", data["tableau_15"])

    #     # Fill ZoneTexte 1
    #     if "zonetexte_1" in data:
    #         self._set_text("ZoneTexte 1", data["zonetexte_1"])

    #     # Fill ZoneTexte 12
    #     if "zonetexte_12" in data:
    #         self._set_text("ZoneTexte 12", data["zonetexte_12"])

    #     # Fill ZoneTexte 16
    #     if "zonetexte_16" in data:
    #         self._set_text("ZoneTexte 16", data["zonetexte_16"])

    #     # Fill ZoneTexte 20
    #     if "zonetexte_20" in data:
    #         self._set_text("ZoneTexte 20", data["zonetexte_20"])

    #     # Fill ZoneTexte 5
    #     if "zonetexte_5" in data:
    #         self._set_text("ZoneTexte 5", data["zonetexte_5"])

    #     # Fill ZoneTexte 26
    #     if "zonetexte_26" in data:
    #         self._set_text("ZoneTexte 26", data["zonetexte_26"])

    #     # Fill ZoneTexte 28
    #     if "zonetexte_28" in data:
    #         self._set_text("ZoneTexte 28", data["zonetexte_28"])

    #     # Fill Ellipse 72
    #     if "ellipse_72" in data:
    #         self._set_text("Ellipse 72", data["ellipse_72"])

    #     # Fill Ellipse 74
    #     if "ellipse_74" in data:
    #         self._set_text("Ellipse 74", data["ellipse_74"])

    #     # Fill Ellipse 75
    #     if "ellipse_75" in data:
    #         self._set_text("Ellipse 75", data["ellipse_75"])

    #     # Fill ZoneTexte 23
    #     if "zonetexte_23" in data:
    #         self._set_text("ZoneTexte 23", data["zonetexte_23"])

    #     # Fill ZoneTexte 76
    #     if "zonetexte_76" in data:
    #         self._set_text("ZoneTexte 76", data["zonetexte_76"])

    #     # Fill ZoneTexte 78
    #     if "zonetexte_78" in data:
    #         self._set_text("ZoneTexte 78", data["zonetexte_78"])

    #     # Fill ZoneTexte 29
    #     if "zonetexte_29" in data:
    #         self._set_text("ZoneTexte 29", data["zonetexte_29"])

    #     # Fill ZoneTexte 30
    #     if "zonetexte_30" in data:
    #         self._set_text("ZoneTexte 30", data["zonetexte_30"])

    #     # Fill ZoneTexte 31
    #     if "zonetexte_31" in data:
    #         self._set_text("ZoneTexte 31", data["zonetexte_31"])

    # def fill_slide_type_157(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #157.

    #     Modifiable elements (4):
    #     - Title 2 [TITLE]
    #     - Rectangle 31
    #     - Rectangle 32
    #     - Rectangle 33

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(157)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 2
    #     if "title" in data:
    #         self._set_text("Title 2", data["title"])

    #     # Fill Rectangle 31
    #     if "rectangle_31" in data:
    #         self._set_text("Rectangle 31", data["rectangle_31"])

    #     # Fill Rectangle 32
    #     if "rectangle_32" in data:
    #         self._set_text("Rectangle 32", data["rectangle_32"])

    #     # Fill Rectangle 33
    #     if "rectangle_33" in data:
    #         self._set_text("Rectangle 33", data["rectangle_33"])

    # def fill_slide_type_158(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #158.

    #     Modifiable elements (4):
    #     - Title 1 [TITLE]
    #     - Rectangle 13
    #     - Rectangle 14
    #     - Rectangle 32

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(158)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill Rectangle 13
    #     if "rectangle_13" in data:
    #         self._set_text("Rectangle 13", data["rectangle_13"])

    #     # Fill Rectangle 14
    #     if "rectangle_14" in data:
    #         self._set_text("Rectangle 14", data["rectangle_14"])

    #     # Fill Rectangle 32
    #     if "rectangle_32" in data:
    #         self._set_text("Rectangle 32", data["rectangle_32"])

    # def fill_slide_type_159(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #159.

    #     Modifiable elements (7):
    #     - Oval 40
    #     - Title 1 [TITLE]
    #     - Rectangle 64
    #     - Rectangle 65
    #     - Rectangle 66
    #     - Rectangle 67
    #     - Rectangle 68

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(159)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Oval 40
    #     if "oval_40" in data:
    #         self._set_text("Oval 40", data["oval_40"])

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill Rectangle 64
    #     if "rectangle_64" in data:
    #         self._set_text("Rectangle 64", data["rectangle_64"])

    #     # Fill Rectangle 65
    #     if "rectangle_65" in data:
    #         self._set_text("Rectangle 65", data["rectangle_65"])

    #     # Fill Rectangle 66
    #     if "rectangle_66" in data:
    #         self._set_text("Rectangle 66", data["rectangle_66"])

    #     # Fill Rectangle 67
    #     if "rectangle_67" in data:
    #         self._set_text("Rectangle 67", data["rectangle_67"])

    #     # Fill Rectangle 68
    #     if "rectangle_68" in data:
    #         self._set_text("Rectangle 68", data["rectangle_68"])

    # def fill_slide_type_160(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #160.

    #     Modifiable elements (17):
    #     - Rectangle 13
    #     - Rectangle 15
    #     - Rectangle 18
    #     - Rectangle 19
    #     - Text Box 23
    #     - Text Box 24
    #     - Text Box 25
    #     - Rectangle 26
    #     - Rectangle 27
    #     - Rectangle 28
    #     - Rectangle 29
    #     - Rectangle 30
    #     - Rectangle 33
    #     - Rectangle 31
    #     - Rectangle 32
    #     - ZoneTexte 7
    #     - Title 2 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(160)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Rectangle 13
    #     if "rectangle_13" in data:
    #         self._set_text("Rectangle 13", data["rectangle_13"])

    #     # Fill Rectangle 15
    #     if "rectangle_15" in data:
    #         self._set_text("Rectangle 15", data["rectangle_15"])

    #     # Fill Rectangle 18
    #     if "rectangle_18" in data:
    #         self._set_text("Rectangle 18", data["rectangle_18"])

    #     # Fill Rectangle 19
    #     if "rectangle_19" in data:
    #         self._set_text("Rectangle 19", data["rectangle_19"])

    #     # Fill Text Box 23
    #     if "text_box_23" in data:
    #         self._set_text("Text Box 23", data["text_box_23"])

    #     # Fill Text Box 24
    #     if "text_box_24" in data:
    #         self._set_text("Text Box 24", data["text_box_24"])

    #     # Fill Text Box 25
    #     if "text_box_25" in data:
    #         self._set_text("Text Box 25", data["text_box_25"])

    #     # Fill Rectangle 26
    #     if "rectangle_26" in data:
    #         self._set_text("Rectangle 26", data["rectangle_26"])

    #     # Fill Rectangle 27
    #     if "rectangle_27" in data:
    #         self._set_text("Rectangle 27", data["rectangle_27"])

    #     # Fill Rectangle 28
    #     if "rectangle_28" in data:
    #         self._set_text("Rectangle 28", data["rectangle_28"])

    #     # Fill Rectangle 29
    #     if "rectangle_29" in data:
    #         self._set_text("Rectangle 29", data["rectangle_29"])

    #     # Fill Rectangle 30
    #     if "rectangle_30" in data:
    #         self._set_text("Rectangle 30", data["rectangle_30"])

    #     # Fill Rectangle 33
    #     if "rectangle_33" in data:
    #         self._set_text("Rectangle 33", data["rectangle_33"])

    #     # Fill Rectangle 31
    #     if "rectangle_31" in data:
    #         self._set_text("Rectangle 31", data["rectangle_31"])

    #     # Fill Rectangle 32
    #     if "rectangle_32" in data:
    #         self._set_text("Rectangle 32", data["rectangle_32"])

    #     # Fill ZoneTexte 7
    #     if "zonetexte_7" in data:
    #         self._set_text("ZoneTexte 7", data["zonetexte_7"])

    #     # Fill Title 2
    #     if "title" in data:
    #         self._set_text("Title 2", data["title"])

    # def fill_slide_type_161(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #161.

    #     Modifiable elements (13):
    #     - Title 20 [TITLE]
    #     - TextBox 21
    #     - TextBox 22
    #     - TextBox 24
    #     - TextBox 25
    #     - TextBox 26
    #     - TextBox 27
    #     - TextBox 28
    #     - TextBox 29
    #     - TextBox 30
    #     - TextBox 34
    #     - TextBox 35
    #     - TextBox 36

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(161)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 20
    #     if "title" in data:
    #         self._set_text("Title 20", data["title"])

    #     # Fill TextBox 21
    #     if "textbox_21" in data:
    #         self._set_text("TextBox 21", data["textbox_21"])

    #     # Fill TextBox 22
    #     if "textbox_22" in data:
    #         self._set_text("TextBox 22", data["textbox_22"])

    #     # Fill TextBox 24
    #     if "textbox_24" in data:
    #         self._set_text("TextBox 24", data["textbox_24"])

    #     # Fill TextBox 25
    #     if "textbox_25" in data:
    #         self._set_text("TextBox 25", data["textbox_25"])

    #     # Fill TextBox 26
    #     if "textbox_26" in data:
    #         self._set_text("TextBox 26", data["textbox_26"])

    #     # Fill TextBox 27
    #     if "textbox_27" in data:
    #         self._set_text("TextBox 27", data["textbox_27"])

    #     # Fill TextBox 28
    #     if "textbox_28" in data:
    #         self._set_text("TextBox 28", data["textbox_28"])

    #     # Fill TextBox 29
    #     if "textbox_29" in data:
    #         self._set_text("TextBox 29", data["textbox_29"])

    #     # Fill TextBox 30
    #     if "textbox_30" in data:
    #         self._set_text("TextBox 30", data["textbox_30"])

    #     # Fill TextBox 34
    #     if "textbox_34" in data:
    #         self._set_text("TextBox 34", data["textbox_34"])

    #     # Fill TextBox 35
    #     if "textbox_35" in data:
    #         self._set_text("TextBox 35", data["textbox_35"])

    #     # Fill TextBox 36
    #     if "textbox_36" in data:
    #         self._set_text("TextBox 36", data["textbox_36"])

    # def fill_slide_type_162(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #162.

    #     Modifiable elements (19):
    #     - Rectangle 3
    #     - Rectangle 4
    #     - Rectangle 5
    #     - Rectangle 6
    #     - Rectangle 7
    #     - Rectangle 8
    #     - Rectangle 9
    #     - Rectangle 10
    #     - Rectangle 13
    #     - Rectangle 14
    #     - Rectangle 15
    #     - Rectangle 16
    #     - Rectangle 17
    #     - Rectangle 18
    #     - Rectangle 21
    #     - Rectangle 22
    #     - Rectangle 443
    #     - Rectangle 444
    #     - Title 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(162)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Rectangle 3
    #     if "rectangle_3" in data:
    #         self._set_text("Rectangle 3", data["rectangle_3"])

    #     # Fill Rectangle 4
    #     if "rectangle_4" in data:
    #         self._set_text("Rectangle 4", data["rectangle_4"])

    #     # Fill Rectangle 5
    #     if "rectangle_5" in data:
    #         self._set_text("Rectangle 5", data["rectangle_5"])

    #     # Fill Rectangle 6
    #     if "rectangle_6" in data:
    #         self._set_text("Rectangle 6", data["rectangle_6"])

    #     # Fill Rectangle 7
    #     if "rectangle_7" in data:
    #         self._set_text("Rectangle 7", data["rectangle_7"])

    #     # Fill Rectangle 8
    #     if "rectangle_8" in data:
    #         self._set_text("Rectangle 8", data["rectangle_8"])

    #     # Fill Rectangle 9
    #     if "rectangle_9" in data:
    #         self._set_text("Rectangle 9", data["rectangle_9"])

    #     # Fill Rectangle 10
    #     if "rectangle_10" in data:
    #         self._set_text("Rectangle 10", data["rectangle_10"])

    #     # Fill Rectangle 13
    #     if "rectangle_13" in data:
    #         self._set_text("Rectangle 13", data["rectangle_13"])

    #     # Fill Rectangle 14
    #     if "rectangle_14" in data:
    #         self._set_text("Rectangle 14", data["rectangle_14"])

    #     # Fill Rectangle 15
    #     if "rectangle_15" in data:
    #         self._set_text("Rectangle 15", data["rectangle_15"])

    #     # Fill Rectangle 16
    #     if "rectangle_16" in data:
    #         self._set_text("Rectangle 16", data["rectangle_16"])

    #     # Fill Rectangle 17
    #     if "rectangle_17" in data:
    #         self._set_text("Rectangle 17", data["rectangle_17"])

    #     # Fill Rectangle 18
    #     if "rectangle_18" in data:
    #         self._set_text("Rectangle 18", data["rectangle_18"])

    #     # Fill Rectangle 21
    #     if "rectangle_21" in data:
    #         self._set_text("Rectangle 21", data["rectangle_21"])

    #     # Fill Rectangle 22
    #     if "rectangle_22" in data:
    #         self._set_text("Rectangle 22", data["rectangle_22"])

    #     # Fill Rectangle 443
    #     if "rectangle_443" in data:
    #         self._set_text("Rectangle 443", data["rectangle_443"])

    #     # Fill Rectangle 444
    #     if "rectangle_444" in data:
    #         self._set_text("Rectangle 444", data["rectangle_444"])

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    # def fill_slide_type_163(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #163.

    #     Modifiable elements (5):
    #     - Titre 1 [TITLE]
    #     - Table 2
    #     - Table 14
    #     - Table 10
    #     - Table 14

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(163)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill Table 2
    #     if "table_2" in data:
    #         self._set_text("Table 2", data["table_2"])

    #     # Fill Table 14
    #     if "table_14" in data:
    #         self._set_text("Table 14", data["table_14"])

    #     # Fill Table 10
    #     if "table_10" in data:
    #         self._set_text("Table 10", data["table_10"])

    #     # Fill Table 14
    #     if "table_14" in data:
    #         self._set_text("Table 14", data["table_14"])

    # def fill_slide_type_164(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #164.

    #     Modifiable elements (9):
    #     - Titre 1 [TITLE]
    #     - Rectangle 5
    #     - Rectangle 6
    #     - Rectangle 7
    #     - Rectangle 8
    #     - ZoneTexte 13
    #     - ZoneTexte 17
    #     - Rectangle 18
    #     - Rectangle 27

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(164)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill Rectangle 5
    #     if "rectangle_5" in data:
    #         self._set_text("Rectangle 5", data["rectangle_5"])

    #     # Fill Rectangle 6
    #     if "rectangle_6" in data:
    #         self._set_text("Rectangle 6", data["rectangle_6"])

    #     # Fill Rectangle 7
    #     if "rectangle_7" in data:
    #         self._set_text("Rectangle 7", data["rectangle_7"])

    #     # Fill Rectangle 8
    #     if "rectangle_8" in data:
    #         self._set_text("Rectangle 8", data["rectangle_8"])

    #     # Fill ZoneTexte 13
    #     if "zonetexte_13" in data:
    #         self._set_text("ZoneTexte 13", data["zonetexte_13"])

    #     # Fill ZoneTexte 17
    #     if "zonetexte_17" in data:
    #         self._set_text("ZoneTexte 17", data["zonetexte_17"])

    #     # Fill Rectangle 18
    #     if "rectangle_18" in data:
    #         self._set_text("Rectangle 18", data["rectangle_18"])

    #     # Fill Rectangle 27
    #     if "rectangle_27" in data:
    #         self._set_text("Rectangle 27", data["rectangle_27"])

    # def fill_slide_type_165(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #165.

    #     Modifiable elements (4):
    #     - Title 6 [TITLE]
    #     - Tableau 12
    #     - ZoneTexte 31
    #     - ZoneTexte 16

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(165)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 6
    #     if "title" in data:
    #         self._set_text("Title 6", data["title"])

    #     # Fill Tableau 12
    #     if "tableau_12" in data:
    #         self._set_text("Tableau 12", data["tableau_12"])

    #     # Fill ZoneTexte 31
    #     if "zonetexte_31" in data:
    #         self._set_text("ZoneTexte 31", data["zonetexte_31"])

    #     # Fill ZoneTexte 16
    #     if "zonetexte_16" in data:
    #         self._set_text("ZoneTexte 16", data["zonetexte_16"])

    # def fill_slide_type_166(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #166.

    #     Modifiable elements (4):
    #     - Titre 1 [TITLE]
    #     - Table 9
    #     - Table 9
    #     - Table 9

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(166)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill Table 9
    #     if "table_9" in data:
    #         self._set_text("Table 9", data["table_9"])

    #     # Fill Table 9
    #     if "table_9" in data:
    #         self._set_text("Table 9", data["table_9"])

    #     # Fill Table 9
    #     if "table_9" in data:
    #         self._set_text("Table 9", data["table_9"])

    # def fill_slide_type_167(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #167.

    #     Modifiable elements (5):
    #     - Titre 1 [TITLE]
    #     - Table 9
    #     - Table 9
    #     - Table 9
    #     - Table 9

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(167)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill Table 9
    #     if "table_9" in data:
    #         self._set_text("Table 9", data["table_9"])

    #     # Fill Table 9
    #     if "table_9" in data:
    #         self._set_text("Table 9", data["table_9"])

    #     # Fill Table 9
    #     if "table_9" in data:
    #         self._set_text("Table 9", data["table_9"])

    #     # Fill Table 9
    #     if "table_9" in data:
    #         self._set_text("Table 9", data["table_9"])

    # def fill_slide_type_168(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #168.

    #     Modifiable elements (8):
    #     - TextBox 32
    #     - Rectangle 40
    #     - TextBox 33
    #     - Rectangle 39
    #     - TextBox 34
    #     - Rectangle 38
    #     - Title 1 [TITLE]
    #     - TextBox 47

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(168)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill TextBox 32
    #     if "textbox_32" in data:
    #         self._set_text("TextBox 32", data["textbox_32"])

    #     # Fill Rectangle 40
    #     if "rectangle_40" in data:
    #         self._set_text("Rectangle 40", data["rectangle_40"])

    #     # Fill TextBox 33
    #     if "textbox_33" in data:
    #         self._set_text("TextBox 33", data["textbox_33"])

    #     # Fill Rectangle 39
    #     if "rectangle_39" in data:
    #         self._set_text("Rectangle 39", data["rectangle_39"])

    #     # Fill TextBox 34
    #     if "textbox_34" in data:
    #         self._set_text("TextBox 34", data["textbox_34"])

    #     # Fill Rectangle 38
    #     if "rectangle_38" in data:
    #         self._set_text("Rectangle 38", data["rectangle_38"])

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill TextBox 47
    #     if "textbox_47" in data:
    #         self._set_text("TextBox 47", data["textbox_47"])

    # def fill_slide_type_169(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #169.

    #     Modifiable elements (2):
    #     - Title 1 [TITLE]
    #     - TextBox 47

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(169)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill TextBox 47
    #     if "textbox_47" in data:
    #         self._set_text("TextBox 47", data["textbox_47"])

    # def fill_slide_type_170(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #170.

    #     Modifiable elements (2):
    #     - Rectangle 3
    #     - Title 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(170)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Rectangle 3
    #     if "rectangle_3" in data:
    #         self._set_text("Rectangle 3", data["rectangle_3"])

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    # def fill_slide_type_171(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #171.

    #     Modifiable elements (14):
    #     - Rectangle 3
    #     - Title 1 [TITLE]
    #     - TextBox 6
    #     - TextBox 7
    #     - TextBox 8
    #     - TextBox 14
    #     - TextBox 34
    #     - TextBox 36
    #     - Rectangle 2
    #     - Rectangle 5
    #     - Rectangle 19
    #     - Rectangle 20
    #     - Rectangle 26
    #     - Rectangle 29

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(171)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Rectangle 3
    #     if "rectangle_3" in data:
    #         self._set_text("Rectangle 3", data["rectangle_3"])

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill TextBox 6
    #     if "textbox_6" in data:
    #         self._set_text("TextBox 6", data["textbox_6"])

    #     # Fill TextBox 7
    #     if "textbox_7" in data:
    #         self._set_text("TextBox 7", data["textbox_7"])

    #     # Fill TextBox 8
    #     if "textbox_8" in data:
    #         self._set_text("TextBox 8", data["textbox_8"])

    #     # Fill TextBox 14
    #     if "textbox_14" in data:
    #         self._set_text("TextBox 14", data["textbox_14"])

    #     # Fill TextBox 34
    #     if "textbox_34" in data:
    #         self._set_text("TextBox 34", data["textbox_34"])

    #     # Fill TextBox 36
    #     if "textbox_36" in data:
    #         self._set_text("TextBox 36", data["textbox_36"])

    #     # Fill Rectangle 2
    #     if "rectangle_2" in data:
    #         self._set_text("Rectangle 2", data["rectangle_2"])

    #     # Fill Rectangle 5
    #     if "rectangle_5" in data:
    #         self._set_text("Rectangle 5", data["rectangle_5"])

    #     # Fill Rectangle 19
    #     if "rectangle_19" in data:
    #         self._set_text("Rectangle 19", data["rectangle_19"])

    #     # Fill Rectangle 20
    #     if "rectangle_20" in data:
    #         self._set_text("Rectangle 20", data["rectangle_20"])

    #     # Fill Rectangle 26
    #     if "rectangle_26" in data:
    #         self._set_text("Rectangle 26", data["rectangle_26"])

    #     # Fill Rectangle 29
    #     if "rectangle_29" in data:
    #         self._set_text("Rectangle 29", data["rectangle_29"])

    # def fill_slide_type_172(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #172.

    #     Modifiable elements (10):
    #     - Title 1 [TITLE]
    #     - Pentagon 2
    #     - Pentagon 3
    #     - Pentagon 4
    #     - Pentagon 5
    #     - Rectangle 10
    #     - Rectangle 11
    #     - Rectangle 12
    #     - Rectangle 13
    #     - Content Placeholder 2

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(172)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill Pentagon 2
    #     if "pentagon_2" in data:
    #         self._set_text("Pentagon 2", data["pentagon_2"])

    #     # Fill Pentagon 3
    #     if "pentagon_3" in data:
    #         self._set_text("Pentagon 3", data["pentagon_3"])

    #     # Fill Pentagon 4
    #     if "pentagon_4" in data:
    #         self._set_text("Pentagon 4", data["pentagon_4"])

    #     # Fill Pentagon 5
    #     if "pentagon_5" in data:
    #         self._set_text("Pentagon 5", data["pentagon_5"])

    #     # Fill Rectangle 10
    #     if "rectangle_10" in data:
    #         self._set_text("Rectangle 10", data["rectangle_10"])

    #     # Fill Rectangle 11
    #     if "rectangle_11" in data:
    #         self._set_text("Rectangle 11", data["rectangle_11"])

    #     # Fill Rectangle 12
    #     if "rectangle_12" in data:
    #         self._set_text("Rectangle 12", data["rectangle_12"])

    #     # Fill Rectangle 13
    #     if "rectangle_13" in data:
    #         self._set_text("Rectangle 13", data["rectangle_13"])

    #     # Fill Content Placeholder 2
    #     if "content_placeholder_2" in data:
    #         self._set_text("Content Placeholder 2",
    #                        data["content_placeholder_2"])

    # def fill_slide_type_173(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #173.

    #     Modifiable elements (2):
    #     - Title 1 [TITLE]
    #     - Content Placeholder 2

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(173)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill Content Placeholder 2
    #     if "content_placeholder_2" in data:
    #         self._set_text("Content Placeholder 2",
    #                        data["content_placeholder_2"])

    # def fill_slide_type_174(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #174.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(174)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    # def fill_slide_type_175(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #175.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(175)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    # def fill_slide_type_176(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #176.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(176)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    # def fill_slide_type_177(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #177.

    #     Modifiable elements (9):
    #     - Titre 1 [TITLE]
    #     - TextBox 39
    #     - TextBox 40
    #     - TextBox 42
    #     - TextBox 43
    #     - TextBox 45
    #     - TextBox 46
    #     - TextBox 48
    #     - TextBox 49

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(177)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill TextBox 39
    #     if "textbox_39" in data:
    #         self._set_text("TextBox 39", data["textbox_39"])

    #     # Fill TextBox 40
    #     if "textbox_40" in data:
    #         self._set_text("TextBox 40", data["textbox_40"])

    #     # Fill TextBox 42
    #     if "textbox_42" in data:
    #         self._set_text("TextBox 42", data["textbox_42"])

    #     # Fill TextBox 43
    #     if "textbox_43" in data:
    #         self._set_text("TextBox 43", data["textbox_43"])

    #     # Fill TextBox 45
    #     if "textbox_45" in data:
    #         self._set_text("TextBox 45", data["textbox_45"])

    #     # Fill TextBox 46
    #     if "textbox_46" in data:
    #         self._set_text("TextBox 46", data["textbox_46"])

    #     # Fill TextBox 48
    #     if "textbox_48" in data:
    #         self._set_text("TextBox 48", data["textbox_48"])

    #     # Fill TextBox 49
    #     if "textbox_49" in data:
    #         self._set_text("TextBox 49", data["textbox_49"])

    # def fill_slide_type_178(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #178.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(178)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    # def fill_slide_type_179(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #179.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(179)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    # def fill_slide_type_180(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #180.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(180)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    # def fill_slide_type_181(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #181.

    #     Modifiable elements (8):
    #     - Title 1 [TITLE]
    #     - Rectangle 6
    #     - Rectangle 10
    #     - Rectangle 14
    #     - TextBox 22
    #     - Rectangle 2
    #     - Rectangle 9
    #     - Rectangle 12

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(181)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill Rectangle 6
    #     if "rectangle_6" in data:
    #         self._set_text("Rectangle 6", data["rectangle_6"])

    #     # Fill Rectangle 10
    #     if "rectangle_10" in data:
    #         self._set_text("Rectangle 10", data["rectangle_10"])

    #     # Fill Rectangle 14
    #     if "rectangle_14" in data:
    #         self._set_text("Rectangle 14", data["rectangle_14"])

    #     # Fill TextBox 22
    #     if "textbox_22" in data:
    #         self._set_text("TextBox 22", data["textbox_22"])

    #     # Fill Rectangle 2
    #     if "rectangle_2" in data:
    #         self._set_text("Rectangle 2", data["rectangle_2"])

    #     # Fill Rectangle 9
    #     if "rectangle_9" in data:
    #         self._set_text("Rectangle 9", data["rectangle_9"])

    #     # Fill Rectangle 12
    #     if "rectangle_12" in data:
    #         self._set_text("Rectangle 12", data["rectangle_12"])

    # def fill_slide_type_182(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #182.

    #     Modifiable elements (2):
    #     - Title 1 [TITLE]
    #     - TextBox 22

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(182)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    #     # Fill TextBox 22
    #     if "textbox_22" in data:
    #         self._set_text("TextBox 22", data["textbox_22"])

    # def fill_slide_type_183(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #183.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(183)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    # def fill_slide_type_184(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #184.

    #     Modifiable elements (1):
    #     - Title 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(184)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    # def fill_slide_type_185(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #185.

    #     Modifiable elements (1):
    #     - Title 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(185)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Title 1
    #     if "title" in data:
    #         self._set_text("Title 1", data["title"])

    # def fill_slide_type_186(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #186.

    #     Modifiable elements (9):
    #     - Titre 1 [TITLE]
    #     - Rectangle 2
    #     - Rectangle 3
    #     - Rectangle 4
    #     - Rectangle 5
    #     - Rectangle 12
    #     - Rectangle 13
    #     - Rectangle 14
    #     - Rectangle 15

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(186)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill Rectangle 2
    #     if "rectangle_2" in data:
    #         self._set_text("Rectangle 2", data["rectangle_2"])

    #     # Fill Rectangle 3
    #     if "rectangle_3" in data:
    #         self._set_text("Rectangle 3", data["rectangle_3"])

    #     # Fill Rectangle 4
    #     if "rectangle_4" in data:
    #         self._set_text("Rectangle 4", data["rectangle_4"])

    #     # Fill Rectangle 5
    #     if "rectangle_5" in data:
    #         self._set_text("Rectangle 5", data["rectangle_5"])

    #     # Fill Rectangle 12
    #     if "rectangle_12" in data:
    #         self._set_text("Rectangle 12", data["rectangle_12"])

    #     # Fill Rectangle 13
    #     if "rectangle_13" in data:
    #         self._set_text("Rectangle 13", data["rectangle_13"])

    #     # Fill Rectangle 14
    #     if "rectangle_14" in data:
    #         self._set_text("Rectangle 14", data["rectangle_14"])

    #     # Fill Rectangle 15
    #     if "rectangle_15" in data:
    #         self._set_text("Rectangle 15", data["rectangle_15"])

    # def fill_slide_type_187(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #187.

    #     Modifiable elements (11):
    #     - Titre 1 [TITLE]
    #     - Rectangle 2
    #     - Rectangle 3
    #     - Rectangle 4
    #     - Rectangle 5
    #     - Rectangle 6
    #     - Rectangle 12
    #     - Rectangle 13
    #     - Rectangle 14
    #     - Rectangle 15
    #     - Rectangle 16

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(187)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])

    #     # Fill Rectangle 2
    #     if "rectangle_2" in data:
    #         self._set_text("Rectangle 2", data["rectangle_2"])

    #     # Fill Rectangle 3
    #     if "rectangle_3" in data:
    #         self._set_text("Rectangle 3", data["rectangle_3"])

    #     # Fill Rectangle 4
    #     if "rectangle_4" in data:
    #         self._set_text("Rectangle 4", data["rectangle_4"])

    #     # Fill Rectangle 5
    #     if "rectangle_5" in data:
    #         self._set_text("Rectangle 5", data["rectangle_5"])

    #     # Fill Rectangle 6
    #     if "rectangle_6" in data:
    #         self._set_text("Rectangle 6", data["rectangle_6"])

    #     # Fill Rectangle 12
    #     if "rectangle_12" in data:
    #         self._set_text("Rectangle 12", data["rectangle_12"])

    #     # Fill Rectangle 13
    #     if "rectangle_13" in data:
    #         self._set_text("Rectangle 13", data["rectangle_13"])

    #     # Fill Rectangle 14
    #     if "rectangle_14" in data:
    #         self._set_text("Rectangle 14", data["rectangle_14"])

    #     # Fill Rectangle 15
    #     if "rectangle_15" in data:
    #         self._set_text("Rectangle 15", data["rectangle_15"])

    #     # Fill Rectangle 16
    #     if "rectangle_16" in data:
    #         self._set_text("Rectangle 16", data["rectangle_16"])

    # def fill_slide_type_188(self, index:int,  title: str, content: Dict[str, str]):
    #     """
    #     Fill slide template #188.

    #     Modifiable elements (1):
    #     - Titre 1 [TITLE]

    #     Args:
    #         content: JSON string with content to fill
    #     """
    #     # Set the current slide and rebuild shape cache
    #     self.slide = self.presentation.Slides(188)
    #     self._build_shape_cache()

    #     # Parse content
    #     data = json.loads(content)

    #     # Fill Titre 1
    #     if "title" in data:
    #         self._set_text("Titre 1", data["title"])
