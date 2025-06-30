# This file contains data models used throughout the application, defining structures for various data entities.

class Attachment:
    def __init__(self, company, measured, height_in_inches):
        self.company = company
        self.measured = measured
        self.height_in_inches = height_in_inches

class Pole:
    def __init__(self, scid, address, attachments=None):
        self.scid = scid
        self.address = address
        self.attachments = attachments if attachments is not None else []

class Route:
    def __init__(self, line_number, poles, connections):
        self.line_number = line_number
        self.poles = poles
        self.connections = connections

class Configuration:
    def __init__(self, power_company, telecom_providers, power_keywords, telecom_keywords, output_settings, column_mappings):
        self.power_company = power_company
        self.telecom_providers = telecom_providers
        self.power_keywords = power_keywords
        self.telecom_keywords = telecom_keywords
        self.output_settings = output_settings
        self.column_mappings = column_mappings