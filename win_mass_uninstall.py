import os
import sys
import argparse
import logging
import log_helper
import registry_helper

logger = log_helper.setup_logger(name="windows_uninstall", level=logging.DEBUG, log_to_file=False)


def environment_value(environment_name):
    """
    :param environment_name: Name of the environment variable
    :return: Value of the environment variable or the empty string if not exists
    """
    try:
        return os.environ[environment_name]
    except KeyError:
        return ''


class WinUninstallApplication:
    """
    Application for uninstalling Windows applications based on the applications list
    """

    # Necessary commands
    WMIC_ENUMERATE_COMMAND = 'wmic product get name'
    WMIC_UNINSTALL_COMMAND = 'wmic product where name="{0}" call uninstall'

    def __init__(self):
        """
        Read installed Windows applications from:
        * HKEY_CURRENT_USER\\Software\\Microsoft\\Installer\\Products
        * HKEY_CLASSES_ROOT\\Installer\\Products
        Check if application requested to uninstall matches the list
        """
        unique_applications = set()
        unique_applications.update(set(WinUninstallApplication.__read_registry_applications(
            key_hive="HKEY_CURRENT_USER",
            key_path="Software\\Microsoft\\Installer\\Products")))
        unique_applications.update(set(WinUninstallApplication.__read_registry_applications(
            key_hive="HKEY_CLASSES_ROOT",
            key_path="Installer\\Products")))

        self.known_applications = list(unique_applications)
        logger.info("Windows has %d installed applications" % len(self.known_applications))

    def enumerate_installed_applications(self):
        """
        :return: List of installed Windows applications, include those not listed in "Add or Remove Programs"
        Non-Latin symbols are not recognized
        """
        registry_applications = WinUninstallApplication.__read_file_applications()
        for app in registry_applications:
            if app not in self.known_applications:
                logger.warning("Application \"%s\" is not in expected registry key" % app)
        return registry_applications

    def perform_uninstall(self, app):
        """
        Uninstall single Windows application
        :param app: Application name (must be Latin alphabet)
        """
        if app not in self.known_applications:
            logger.warning("Application \"%s\" is not installed or not recognized" % app)
            return
        logger.info("Uninstall {0}".format(app))
        os.system(WinUninstallApplication.WMIC_UNINSTALL_COMMAND.format(app))

    @staticmethod
    def read_from_file(file_name):
        """
        :param file_name: File name, contains list of application to be uninstalled
        :return: List of lines, without proper check if the application installed
        The check is performed in perform_uninstall() method
        """
        logger.info("Read application contents from file {0}".format(file_name))
        with open(file_name) as f:
            applications = f.readlines()
        logger.info(applications)
        logger.info("%d applications to uninstall" % len(applications))
        result_uninstall_list = [x.strip() for x in applications]
        return result_uninstall_list

    @staticmethod
    def __read_file_applications():
        """
        Read installed applications string from WMIC output
        WMIC resulted strings should be stripped of end-lines and trailing spaces
        :return: List of installed applications in alphabetical order
        """
        logger.info("Enumerate installed applications. Patience please, it takes some time...")
        wmic_output = os.popen(WinUninstallApplication.WMIC_ENUMERATE_COMMAND).read()
        # Cut all /r, /n and space symbols
        applications_list = wmic_output.rstrip().split('\n\n')
        applications_list = [app.strip() for app in applications_list if len(app)]
        # First item is column name
        applications_list = applications_list[1:]
        applications_list.sort()
        logger.info("Enumerate finished with {0} installed applications".format(len(applications_list)))
        return applications_list

    @staticmethod
    def __read_registry_applications(key_hive, key_path):
        """
        :return: List of known Windows applications list, located in HKEY_CLASSES_ROOT\Installer\Products
        """
        products = registry_helper.enumerate_key_subkeys(key_hive, key_path)
        if len(products) == 0:
            logger.error("%s\\%s is unexpectedly empty" % key_hive, key_path)
            sys.exit(0)
        product_names = []
        for product in products:
            product_name = registry_helper.read_value(key_hive=key_hive,
                                                      key_path="{0}\\{1}".format(key_path, product),
                                                      value_name="ProductName")
            product_names.append(product_name)
        product_names = [product[0] for product in product_names]
        return product_names


def main():
    """
    Uninstall applications based on list, or simply retrreive the list of installed applications
    :return: System return code
    """
    parser = argparse.ArgumentParser(description='Command-line params')
    parser.add_argument('--installed-apps',
                        help='Get the list of installed applications',
                        dest='installed_apps',
                        action='store_true',
                        required=False)
    parser.add_argument('--uninstall',
                        help='Uninstall applications based on text file',
                        dest='uninstall',
                        type=str,
                        required=False)

    args = parser.parse_args()

    windows_apps = WinUninstallApplication()
    if args.installed_apps:
        installed_applications = windows_apps.enumerate_installed_applications()
        for app in installed_applications:
            print(app)
        return 0

    if args.uninstall:
        apps = WinUninstallApplication.read_from_file(args.uninstall)
        for app in apps:
            windows_apps.perform_uninstall(app)

    return 0


###########################################################################
if __name__ == '__main__':
    sys.exit(main())
