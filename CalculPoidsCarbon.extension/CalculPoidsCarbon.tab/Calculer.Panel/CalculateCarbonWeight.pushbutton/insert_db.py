if not element_quantity_dict:
    from main_script import element_quantity_dict, doc, materiaux_inconnus, project_name, phase_nom, final_phase_calcul
from outil import round_3_decimals
import socket
from datetime import date
from sqlalchemy import create_engine, Column, Integer, String, Date, ForeignKey, URL
from sqlalchemy.orm import Session, relationship, declarative_base
from itertools import zip_longest
from Autodesk.Revit.DB import (
    ElementId
)
element_quantity_dict = element_quantity_dict
material_class_data = []
sous_projet_data = []
family_data = []
type_data = []
instance_data = []
total = 0
for key, value in element_quantity_dict.items():
    element_idee = ElementId(int(key))
    element = doc.GetElement(element_idee)
    elements_calculated_for_phase = len(element_quantity_dict)
    materials = element.GetMaterialIds(False)
    total += value
    try:
        sous_projet_name = element.LookupParameter("Sous-projet").AsValueString()
        if all(sous_projet_name not in d.keys() for d in sous_projet_data):
            sous_projet_data.append(
                {
                    sous_projet_name: {
                        "carbon_weight": 0,
                        "instance_calculated": 0,
                        "material_class_calculated": [],
                        "family_calculated": [],
                        "type_calculated": [],
                    }
                }
            )
        for sous_projet in sous_projet_data:
            if sous_projet_name in sous_projet.keys():
                # Update "carbon_weight" for the matching sous_projet_name
                sous_projet[sous_projet_name]["carbon_weight"] += value
                sous_projet[sous_projet_name]["instance_calculated"] += 1
                if (
                    mat_class
                    not in sous_projet[sous_projet_name][
                        "material_class_calculated"
                    ]
                ):
                    sous_projet[sous_projet_name][
                        "material_class_calculated"
                    ].append(mat_class)
                if (
                    type_name
                    not in sous_projet[sous_projet_name]["type_calculated"]
                ):
                    sous_projet[sous_projet_name]["type_calculated"].append(
                        type_name
                    )
                if (
                    family_name
                    not in sous_projet[sous_projet_name]["family_calculated"]
                ):
                    sous_projet[sous_projet_name]["family_calculated"].append(
                        family_name
                    )

                # Update material_classes_calculated with the count of material classes
                sous_projet[sous_projet_name]["material_class_count"] = len(
                    set(sous_projet[sous_projet_name]["material_class_calculated"])
                )
                sous_projet[sous_projet_name]["type_count"] = len(
                    set(sous_projet[sous_projet_name]["type_calculated"])
                )
                sous_projet[sous_projet_name]["family_count"] = len(
                    set(sous_projet[sous_projet_name]["family_calculated"])
                )

    except:
        pass

    try:
        family_name = element.Category.Name
        if all(family_name not in d.keys() for d in family_data):
            family_data.append(
                {
                    family_name: {
                        "carbon_weight": 0,
                        "instance_calculated": 0,
                        "material_class_calculated": [],
                        "type_calculated": [],
                    }
                }
            )
        for element_category in family_data:
            if family_name in element_category.keys():
                element_category[family_name]["carbon_weight"] += value
                element_category[family_name]["instance_calculated"] += 1
                if (
                    mat_class
                    not in element_category[family_name][
                        "material_class_calculated"
                    ]
                ):
                    element_category[family_name][
                        "material_class_calculated"
                    ].append(mat_class)
                if (
                    type_name
                    not in element_category[family_name]["type_calculated"]
                ):
                    element_category[family_name]["type_calculated"].append(
                        type_name
                    )

                # Update material_classes_calculated with the count of material classes
                element_category[family_name]["material_class_count"] = len(
                    set(element_category[family_name]["material_class_calculated"])
                )
                element_category[family_name]["type_count"] = len(
                    set(element_category[family_name]["type_calculated"])
                )
    except:
        pass

    try:
        type_name = element.LookupParameter("Type").AsValueString()
        if all(type_name not in d.keys() for d in type_data):
            type_data.append(
                {
                    type_name: {
                        "carbon_weight": 0,
                        "instance_calculated": 0,
                        "material_class_calculated": [],
                        "family": family_name,
                    }
                }
            )
        for type in type_data:
            if type_name in type.keys():
                # Update "carbon_weight" for the matching type_name
                type[type_name]["carbon_weight"] += value
                type[type_name]["instance_calculated"] += 1
                if mat_class not in type[type_name]["material_class_calculated"]:
                    type[type_name]["material_class_calculated"].append(mat_class)

                # Update material_classes_calculated with the count of material classes
                type[type_name]["material_class_count"] = len(
                    set(type[type_name]["material_class_calculated"])
                )
    except:
        pass
    multiple_materials = ""
    try:
        type_name = element.LookupParameter("Type").AsValueString()
        family_name = element.Category.Name
        sous_projet_name = element.LookupParameter("Sous-projet").AsValueString()
        mat_class = elem_mat.MaterialClass
        # if len(materials) > 1:
        #     for m in materials:
        #         elem_m = doc.GetElement(m)
        #         m_class = elem_m.MaterialClass
        #         multiple_materials += str(m_class + ", ")
        if all(key not in d.keys() for d in instance_data) and len(materials) == 1:
            instance_data.append(
                {
                    key: {
                        "carbon_weight": value,
                        "family": family_name,
                        "sous_projet": sous_projet_name,
                        "type": type_name,
                        "material_class": mat_class,
                    }
                }
            )
        else:
            instance_data.append(
                {
                    key: {
                        "carbon_weight": value,
                        "family": family_name,
                        "sous_projet": sous_projet_name,
                        "type": type_name,
                        "material_class": "multiple",
                    }
                }
            )
    except Exception as e:
        print(e)
    for material in materials:
        elem_mat = doc.GetElement(material)
        mat_class = elem_mat.MaterialClass
        if (
            all(mat_class not in d.keys() for d in material_class_data)
            and mat_class not in materiaux_inconnus
        ):
            material_class_data.append(
                {
                    mat_class: {
                        "carbon_weight": 0,
                        "instance_calculated": 0,
                    }
                }
            )
        for mat in material_class_data:
            if mat_class in mat.keys():
                # Update "carbon_weight" for the matching type_name
                mat[mat_class]["carbon_weight"] += value
                mat[mat_class]["instance_calculated"] += 1
    material_classes_for_phase = len(material_class_data)

family_data_for_phase = len(family_data)
type_data_for_phase = len(type_data)
project_data = project_name
phase_data = (
    phase_nom,
    round_3_decimals(total),
    date.today(),
    elements_calculated_for_phase,
    material_classes_for_phase,
    family_data_for_phase,
    type_data_for_phase,
)

ip_address = socket.gethostbyname("PO-ARP-6155")

if final_phase_calcul:
    Base = declarative_base()

    class Project(Base):
        __tablename__ = "project"

        project_id = Column(Integer, primary_key=True)
        project_name = Column(String)

    class Phase(Base):
        __tablename__ = "phase"

        phase_id = Column(Integer, primary_key=True)
        phase_name = Column(String)
        carbon_weight_kg = Column(Integer)
        date_calculated = Column(Date)
        project_id = Column(Integer, ForeignKey("project.project_id"))
        instance_calculated = Column(Integer)
        material_class_calculated = Column(Integer)
        family_calculated = Column(Integer)
        type_calculated = Column(Integer)

        # Define a relationship to the Project class
        project = relationship("Project", back_populates="phase")

    Project.phase = relationship(
        "Phase", order_by=Phase.phase_id, back_populates="project"
    )

    class SousProjet(Base):
        __tablename__ = "sous_projet"

        sous_projet_id = Column(Integer, primary_key=True)
        sous_projet_name = Column(String)
        carbon_weight_kg = Column(Integer)
        project_id = Column(Integer, ForeignKey("project.project_id"))
        phase_id = Column(Integer, ForeignKey("phase.phase_id"))
        instance_calculated = Column(Integer)
        material_class_calculated = Column(Integer)
        family_calculated = Column(Integer)
        type_calculated = Column(Integer)

        # Define relationships to the Project and Phase classes
        project = relationship("Project", back_populates="sous_projet")
        phase = relationship("Phase", back_populates="sous_projet")

    # Add the sous_projets relationship to the Project and Phase classes
    Project.sous_projet = relationship(
        "SousProjet", order_by=SousProjet.sous_projet_id, back_populates="project"
    )
    Phase.sous_projet = relationship(
        "SousProjet", order_by=SousProjet.sous_projet_id, back_populates="phase"
    )

    class Family(Base):
        __tablename__ = "family"

        family_id = Column(Integer, primary_key=True)
        family_name = Column(String)
        carbon_weight_kg = Column(Integer)
        project_id = Column(Integer, ForeignKey("project.project_id"))
        phase_id = Column(Integer, ForeignKey("phase.phase_id"))
        instance_calculated = Column(Integer)
        material_class_calculated = Column(Integer)
        type_calculated = Column(Integer)

        # Define relationships to the Project, Phase, and Sous_projet classes
        project = relationship("Project", back_populates="family")
        phase = relationship("Phase", back_populates="family")

    # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
    Project.family = relationship(
        "Family", order_by=Family.family_id, back_populates="project"
    )
    Phase.family = relationship(
        "Family", order_by=Family.family_id, back_populates="phase"
    )

    class Type(Base):
        __tablename__ = "type"

        type_id = Column(Integer, primary_key=True)
        type_name = Column(String)
        carbon_weight_kg = Column(Integer)
        project_id = Column(Integer, ForeignKey("project.project_id"))
        phase_id = Column(Integer, ForeignKey("phase.phase_id"))
        family = Column(String)
        instance_calculated = Column(Integer)
        material_class_calculated = Column(Integer)

        # Define relationships to the Project, Phase, and Sous_projet classes
        project = relationship("Project", back_populates="type")
        phase = relationship("Phase", back_populates="type")

    # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
    Project.type = relationship(
        "Type", order_by=Type.type_id, back_populates="project"
    )
    Phase.type = relationship("Type", order_by=Type.type_id, back_populates="phase")

    class Instance(Base):
        __tablename__ = "instance"

        instance_id = Column(Integer, primary_key=True)
        carbon_weight_kg = Column(Integer)
        project_id = Column(Integer, ForeignKey("project.project_id"))
        phase_id = Column(Integer, ForeignKey("phase.phase_id"))
        material_class = Column(String)
        sous_projet = Column(String)
        family = Column(String)
        type = Column(String)

        # Define relationships to the Project, Phase, and Sous_projet classes
        project = relationship("Project", back_populates="instance")
        phase = relationship("Phase", back_populates="instance")

    # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
    Project.instance = relationship(
        "Instance", order_by=Instance.instance_id, back_populates="project"
    )
    Phase.instance = relationship(
        "Instance", order_by=Instance.instance_id, back_populates="phase"
    )

    class MaterialClass(Base):
        __tablename__ = "material_class"

        material_class_id = Column(Integer, primary_key=True)
        material_class_name = Column(String)
        carbon_weight_kg = Column(Integer)
        instance_calculated = Column(Integer)
        project_id = Column(Integer, ForeignKey("project.project_id"))
        phase_id = Column(Integer, ForeignKey("phase.phase_id"))

        # Define relationships to the Project, Phase, and Sous_projet classes
        project = relationship("Project", back_populates="material_class")
        phase = relationship("Phase", back_populates="material_class")

    # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
    Project.material_class = relationship(
        "MaterialClass",
        order_by=MaterialClass.material_class_id,
        back_populates="project",
    )
    Phase.material_class = relationship(
        "MaterialClass",
        order_by=MaterialClass.material_class_id,
        back_populates="phase",
    )

    def insert_data_into_database(
        project_data,
        phase_data,
        sous_projet_data,
        family_data,
        type_data,
        instance_data,
        material_class_data,
    ):
        """
        Inserts data into the specified database tables.

        Parameters:
        - project_data (str): Name of the project.
        - phase_data (tuple): Tuple containing phase data in the following format:
            (phase_name, carbon_weight_kg, instance_calculated, material_class_calculated,
            family_calculated, type_calculated).
        - sous_projet_data (list): List of dictionaries containing sous-projet data.
        - family_data (list): List of dictionaries containing family data.
        - type_data (list): List of dictionaries containing type data.
        - instance_data (list): List of dictionaries containing instance data.
        - material_class_data (list): List of dictionaries containing material class data.
        """
        # Set your database connection details
        url_object = URL.create(
            "postgresql+psycopg2",
            username="postgres",
            password="postgres",  # plain (unescaped) text
            host=ip_address,
            database="carbonbim",
        )
        engine = create_engine(url_object)

        # Create a session to interact with the database
        session = Session(engine)

        # try:
        # Insert data into the 'project' table
        project = Project(
            project_name=project_data,
        )
        session.add(project)

        # Insert data into the 'phases' table
        phase = Phase(
            phase_name=phase_data[0],
            carbon_weight_kg=phase_data[1],
            date_calculated=date.today(),
            instance_calculated=phase_data[3],
            material_class_calculated=phase_data[4],
            family_calculated=phase_data[5],
            type_calculated=phase_data[6],
        )
        project.phase.append(phase)

        session.add(phase)

        # Iterate through sous_projet_data, material_class_data, and family_data, etc. simultaneously
        for (
            instance_entry,
            sous_projet_entry,
            family_entry,
            type_entry,
            material_class_entry,
        ) in zip_longest(
            instance_data,
            sous_projet_data,
            family_data,
            type_data,
            material_class_data,
            fillvalue=None,
        ):
            # Use the unique identifier as sous_projet_name, material_class_name, and element_name
            if instance_entry is not None:
                key = list(instance_entry.keys())[0]
            if sous_projet_entry is not None:
                sous_projet_name = list(sous_projet_entry.keys())[0]
            if family_entry is not None:
                family_name = list(family_entry.keys())[0]
            if type_entry is not None:
                type_name = list(type_entry.keys())[0]
            if material_class_entry is not None:
                material_class_name = list(material_class_entry.keys())[0]

            # Insert data into the 'sous_projet' table
            if sous_projet_entry is not None:
                sous_projet = SousProjet(
                    sous_projet_name=sous_projet_name,
                    carbon_weight_kg=sous_projet_entry[sous_projet_name][
                        "carbon_weight"
                    ],
                    instance_calculated=sous_projet_entry[sous_projet_name][
                        "instance_calculated"
                    ],
                    material_class_calculated=sous_projet_entry[sous_projet_name][
                        "material_class_count"
                    ],
                    family_calculated=sous_projet_entry[sous_projet_name][
                        "family_count"
                    ],
                    type_calculated=sous_projet_entry[sous_projet_name][
                        "type_count"
                    ],
                    phase=phase,
                    project=project,
                )
                session.add(sous_projet)

            # Insert data into the 'family' table
            if family_entry is not None:
                family = Family(
                    family_name=family_name,
                    carbon_weight_kg=family_entry[family_name]["carbon_weight"],
                    instance_calculated=family_entry[family_name][
                        "instance_calculated"
                    ],
                    material_class_calculated=family_entry[family_name][
                        "material_class_count"
                    ],
                    type_calculated=family_entry[family_name]["type_count"],
                    phase=phase,
                    project=project,
                )
                session.add(family)

            # Insert data into the 'type' table
            if type_entry is not None:
                type = Type(
                    type_name=type_name,
                    carbon_weight_kg=type_entry[type_name]["carbon_weight"],
                    instance_calculated=type_entry[type_name][
                        "instance_calculated"
                    ],
                    material_class_calculated=type_entry[type_name][
                        "material_class_count"
                    ],
                    family=type_entry[type_name]["family"],
                    phase=phase,
                    project=project,
                )
                session.add(type)

            # Insert data into the 'instance' table
            if instance_entry is not None:
                instance = Instance(
                    carbon_weight_kg=instance_entry[key]["carbon_weight"],
                    family=instance_entry[key]["family"],
                    sous_projet=instance_entry[key]["sous_projet"],
                    type=instance_entry[key]["type"],
                    material_class=instance_entry[key]["material_class"],
                    phase=phase,
                    project=project,
                )
            session.add(instance)
            # Insert data into the 'material_class' table
            if material_class_entry is not None:
                material_class = MaterialClass(
                    material_class_name=material_class_name,
                    carbon_weight_kg=material_class_entry[material_class_name][
                        "carbon_weight"
                    ],
                    instance_calculated=material_class_entry[material_class_name][
                        "instance_calculated"
                    ],
                    phase=phase,
                    project=project,
                )
                session.add(material_class)

        # Commit the transaction
        session.commit()

        print("Data inserted successfully.")

        # except Exception as e:
        #     print(f"Error: {str(e)}")

        # finally:
        #     # Close the session
        session.close()

    insert_data_into_database(
        project_data,
        phase_data,
        sous_projet_data,
        family_data,
        type_data,
        instance_data,
        material_class_data,
    )
