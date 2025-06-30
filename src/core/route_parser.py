import logging
try:
    from .utils import Utils
except ImportError:
    from utils import Utils

class RouteParser:
    """Handles manual route parsing and validation"""
    
    @staticmethod
    def parse_manual_routes(route_text, ignore_keywords=None):
        """Parse manual route definitions"""
        routes = []
        lines = route_text.strip().split('\n')
        
        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            if not line:
                continue
                
            route_segments = line.split(';')
            
            for segment in route_segments:
                segment = segment.strip()
                if not segment:
                    continue
                    
                raw_poles = [pole.strip() for pole in segment.split(',') if pole.strip()]
                poles = [Utils.normalize_scid(pole, ignore_keywords) for pole in raw_poles]
                
                if len(poles) < 2:
                    logging.warning(f"Route line {line_num}: Skipping route with less than 2 poles: {segment}")
                    continue
                
                route_connections = []
                for i in range(len(poles) - 1):
                    route_connections.append((poles[i], poles[i + 1]))
                
                routes.append({
                    'line_number': line_num,
                    'poles': poles,
                    'connections': route_connections
                })
                
                logging.info(f"Parsed route {len(routes)}: {' → '.join(poles)}")
        
        return routes