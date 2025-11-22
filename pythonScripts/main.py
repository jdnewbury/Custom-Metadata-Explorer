# check out the documentation @ https://extentor-au.atlassian.net/wiki/spaces/TQ/pages/2531623189/Automated+Salesforce+Config+Matrix
# to generate the config matrix, run (double click) the main.py executable from the parent folder of PythonScripts

import logging
import utils

def main():
    # Configure application logging
    utils.configure_logging()

    try:
        # Prompt the user for a directory and set up folder and file global vars
        utils.select_project_directory()

        # Open/create the config file
        utils.handle_config_file(utils.config_matrix_path, utils.config_file_name)

        # Run the scripts which generate metadata and create tabs in the config matrix
        import allFieldsScript
        import assignmentRulesScript
        import classesScript
        import connectedAppsScript
        import customFieldsScript
        import customMetadataTypesScript
        import flowsScript
        import installedPackagesScript
        import listViewsScript
        import omniDataTransformScript
        import omniIntegrationProceduresScript
        import omniScriptsScript
        import omniUICardScript
        import permissionsetsScript
        import platformEventsScript
        import profilesScript
        import queuesScript
        import recordTypesScript
        import reportsScript
        import triggersScript
        import validationRulesScript

        # Remove the sheet that is created by default when the config matrix is created
        utils.remove_default_sheet(utils.config_file_path)

        # Open the completed config matrix in Windows
        utils.open_excel_visibly(utils.config_file_path)
        print("The config matrix was generated successfully.")
        input("Press Enter to exit...")


    except Exception as e:
        logging.error(f"An error occurred: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()
