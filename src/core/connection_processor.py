import logging
import pandas as pd
from .utils import Utils


class ConnectionProcessor:
    """Handles connection processing logic"""
    
    def __init__(self, qc_reader=None):
        self.qc_reader = qc_reader
    
    def process_connections(self, connections_df, mappings, sections_df):
        """Process connections to generate output rows"""
        result_data = []
        
        # If QC file is active, generate output based on QC connections only
        if self.qc_reader and self.qc_reader.is_active():
            logging.info("QC file is active - filtering output to QC connections only")
            result_data = self._process_qc_filtered_connections(
                connections_df, mappings, sections_df
            )
        else:
            # Process each connection and generate rows (original logic)
            result_data = self._process_standard_connections(
                connections_df, mappings, sections_df
            )
        
        return result_data
    
    def _process_standard_connections(self, connections_df, mappings, sections_df):
        """Process standard connections without QC filtering"""
        result_data = []
        processed_connections = set()
        
        for _, conn in connections_df.iterrows():
            n1, n2 = str(conn['node_id_1']).strip(), str(conn['node_id_2']).strip()
            
            if (n1 in mappings['valid_poles'] and n2 in mappings['valid_poles']):
                connection_key = tuple(sorted([n1, n2]))
                if connection_key not in processed_connections:
                    processed_connections.add(connection_key)
                    
                    scid1 = mappings['node_id_to_scid'][n1]
                    scid2 = mappings['node_id_to_scid'][n2]
                    
                    node1_data = mappings['node_id_to_row'].get(n1, {})
                    node2_data = mappings['node_id_to_row'].get(n2, {})
                    node1_type = str(node1_data.get('node_type', '')).strip().lower()
                    node2_type = str(node2_data.get('node_type', '')).strip().lower()
                    
                    conn_info = {
                        'connection_id': conn.get('connection_id', ''),
                        'span_distance': conn.get('span_distance', '')
                    }
                    
                    # Generate row(s) for this connection
                    row_data = None
                    if node1_type == 'pole' and node2_type == 'reference':
                        # Pole -> Reference: Pole in "Pole" column, Reference in "To Pole" column
                        row_data = (scid1, scid2, conn_info, node1_data)
                    elif node1_type == 'reference' and node2_type == 'pole':
                        # Reference -> Pole: Pole in "Pole" column, Reference in "To Pole" column
                        row_data = (scid2, scid1, conn_info, node2_data)
                    elif node1_type == 'pole' and node2_type == 'pole':
                        # Pole -> Pole: First pole in "Pole" column, Second pole in "To Pole" column
                        row_data = (scid1, scid2, conn_info, node1_data)
                    
                    if row_data:
                        result_data.append(row_data)
        
        return result_data
    
    def _process_qc_filtered_connections(self, connections_df, mappings, sections_df):
        """Process connections when QC file is active"""
        result_data = []
        
        # Get ordered connections from QC file in ORIGINAL format
        qc_original_connections = self.qc_reader.get_original_ordered_connections()
        qc_normalized_connections = self.qc_reader.get_ordered_connections()
        
        logging.info(f"Processing {len(qc_original_connections)} QC connections in specified order")
        logging.info("QC Mode: Using EXACT original Pole and ToPole format from QC file")
        
        # Create lookup for connection data from Excel (bidirectional)
        connection_lookup = {}
        for _, conn in connections_df.iterrows():
            n1, n2 = str(conn['node_id_1']).strip(), str(conn['node_id_2']).strip()
            if n1 in mappings['node_id_to_scid'] and n2 in mappings['node_id_to_scid']:
                scid1 = mappings['node_id_to_scid'][n1]
                scid2 = mappings['node_id_to_scid'][n2]
                
                conn_info = {
                    'connection_id': conn.get('connection_id', ''),
                    'span_distance': conn.get('span_distance', ''),
                    'node1_id': n1,
                    'node2_id': n2
                }
                
                # Store in both directions for lookup
                connection_lookup[(scid1, scid2)] = conn_info
                connection_lookup[(scid2, scid1)] = conn_info
        
        # Process QC connections in the exact order specified in QC file
        for i, (qc_pole_orig, qc_to_pole_orig) in enumerate(qc_original_connections):
            # Get the corresponding normalized versions for data lookup
            qc_pole_norm, qc_to_pole_norm = qc_normalized_connections[i]
            
            # Check if this connection exists in Excel data (try both directions using normalized SCIDs)
            conn_info = None
            if (qc_pole_norm, qc_to_pole_norm) in connection_lookup:
                conn_info = connection_lookup[(qc_pole_norm, qc_to_pole_norm)]
            elif (qc_to_pole_norm, qc_pole_norm) in connection_lookup:
                conn_info = connection_lookup[(qc_to_pole_norm, qc_pole_norm)]
            
            if not conn_info:
                logging.warning(f"QC connection {qc_pole_orig} -> {qc_to_pole_orig} not found in Excel data")
                # Still create a row with available data if SCIDs exist (using normalized for lookup)
                pole_node_data = mappings['scid_to_row'].get(qc_pole_norm, {})
                to_pole_node_data = mappings['scid_to_row'].get(qc_to_pole_norm, {})
                
                if pole_node_data or to_pole_node_data:
                    # Create minimal connection info
                    conn_info = {
                        'connection_id': '',
                        'span_distance': '',
                        'node1_id': '',
                        'node2_id': ''
                    }
                else:
                    continue
            
            # Get node data for the pole specified in QC file (using normalized SCID for lookup)
            pole_node_data = mappings['scid_to_row'].get(qc_pole_norm, {})
            
            # Return QC connection data tuple (original format preserved)
            result_data.append((qc_pole_orig, qc_to_pole_orig, qc_pole_norm, qc_to_pole_norm, conn_info, pole_node_data))
        
        logging.info(f"Generated {len(result_data)} QC-filtered connection tuples")
        return result_data
    
    def build_temp_rows(self, connections_df, mappings, manual_routes, clear_existing_routes):
        """Build temporary rows for processing"""
        temp = {}
        processed = set()
        
        # Initialize all valid poles
        for node_id in mappings['valid_poles']:
            scid = mappings['node_id_to_scid'][node_id]
            node_data = mappings['node_id_to_row'].get(node_id, {})
            guy_info = self._extract_guy_info(node_data.get('mr_note', ''))
            
            temp[scid] = {
                'Pole': scid,
                'Guy Size': '',
                'Guy Lead': ', '.join(guy_info['leads']),
                'Guy Direction': ', '.join(guy_info['directions']),
                'To Pole': '',
                'connection_id': '',
                'span_distance': ''
            }
        
        # Skip Excel connection processing if QC file is active
        if self.qc_reader and self.qc_reader.is_active():
            logging.info("QC file is active - skipping Excel connection processing")
            connection_data = {}
        else:
            # Process Excel connections
            connection_data = self._process_excel_connections(
                connections_df, mappings, temp, processed, clear_existing_routes
            )
        
        # Apply manual routes (only if QC file is not active)
        if manual_routes and not (self.qc_reader and self.qc_reader.is_active()):
            self._apply_manual_routes(manual_routes, temp, connection_data)
        elif manual_routes and self.qc_reader and self.qc_reader.is_active():
            logging.info("QC file is active - manual routes will be ignored in favor of QC connections")
        
        logging.info(f"Built {len(temp)} pole records with routing information")
        return temp
    
    def _process_excel_connections(self, connections_df, mappings, temp, processed, clear_existing_routes):
        """Process connections from Excel data with enhanced reference node logic"""
        logging.info("Processing automatic connections from Excel data...")
        connection_data = {}
        
        for _, conn in connections_df.iterrows():
            n1, n2 = str(conn['node_id_1']).strip(), str(conn['node_id_2']).strip()
            
            if (n1 in mappings['valid_poles'] and n2 in mappings['valid_poles']):
                connection_key = tuple(sorted([n1, n2]))
                if connection_key not in processed:
                    processed.add(connection_key)
                    scid1 = mappings['node_id_to_scid'][n1]
                    scid2 = mappings['node_id_to_scid'][n2]
                    
                    # Get node types to handle reference nodes correctly
                    node1_data = mappings['node_id_to_row'].get(n1, {})
                    node2_data = mappings['node_id_to_row'].get(n2, {})
                    node1_type = str(node1_data.get('node_type', '')).strip().lower()
                    node2_type = str(node2_data.get('node_type', '')).strip().lower()
                    
                    conn_info = {
                        'connection_id': conn.get('connection_id', ''),
                        'span_distance': conn.get('span_distance', '')
                    }
                    
                    # Store connection data for both directions
                    connection_data[(scid1, scid2)] = conn_info
                    connection_data[(scid2, scid1)] = conn_info
                    
                    if not clear_existing_routes:
                        # Handle reference node logic: references must be at 'To Pole'
                        if node2_type == 'reference' and node1_type == 'pole':
                            # scid1 is pole, scid2 is reference
                            temp[scid1].update({'To Pole': scid2, **conn_info})
                        elif node1_type == 'reference' and node2_type == 'pole':
                            # scid2 is pole, scid1 is reference  
                            temp[scid2].update({'To Pole': scid1, **conn_info})
                        elif node1_type == 'pole' and node2_type == 'pole':
                            # Both are poles, use normal connection logic
                            temp[scid1].update({'To Pole': scid2, **conn_info})
                        else:
                            # Default behavior for other cases
                            temp[scid1].update({'To Pole': scid2, **conn_info})
        
        if clear_existing_routes:
            logging.info("Cleared existing route data as requested")
            for scid in temp:
                temp[scid]['To Pole'] = ''
        
        return connection_data
    
    def _apply_manual_routes(self, manual_routes, temp, connection_data):
        """Apply manual route definitions"""
        logging.info(f"Applying {len(manual_routes)} manual routes...")
        
        for route in manual_routes:
            poles = route['poles']
            connections = route['connections']
            
            for from_pole, to_pole in connections:
                if from_pole in temp:
                    # Get connection info if available
                    conn_info = connection_data.get((from_pole, to_pole), {})
                    
                    temp[from_pole].update({
                        'To Pole': to_pole,
                        'connection_id': conn_info.get('connection_id', ''),
                        'span_distance': conn_info.get('span_distance', '')
                    })
                    
                    logging.debug(f"Manual route: {from_pole} -> {to_pole}")
        
        logging.info("Manual routes applied successfully")
    
    def _extract_guy_info(self, note):
        """Extract guy wire information from notes"""
        guy_info = {'leads': [], 'directions': []}
        
        if not note:
            return guy_info
        
        note_lower = str(note).lower()
        
        # Extract guy leads
        import re
        lead_matches = re.findall(r'guy\s*lead[:\s]*([^,\n]+)', note_lower)
        for match in lead_matches:
            clean_lead = match.strip()
            if clean_lead:
                guy_info['leads'].append(clean_lead)
        
        # Extract guy directions
        direction_matches = re.findall(r'guy\s*direction[:\s]*([^,\n]+)', note_lower)
        for match in direction_matches:
            clean_direction = match.strip()
            if clean_direction:
                guy_info['directions'].append(clean_direction)
        
        return guy_info 