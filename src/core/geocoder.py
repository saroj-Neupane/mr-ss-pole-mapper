from pathlib import Path
import csv
import logging
from geopy.geocoders import Nominatim

class Geocoder:
    """Handles geocoding functionality"""
    
    # State name to abbreviation mapping
    STATE_CODES = {
        'alabama': 'AL', 'alaska': 'AK', 'arizona': 'AZ', 'arkansas': 'AR', 'california': 'CA',
        'colorado': 'CO', 'connecticut': 'CT', 'delaware': 'DE', 'florida': 'FL', 'georgia': 'GA',
        'hawaii': 'HI', 'idaho': 'ID', 'illinois': 'IL', 'indiana': 'IN', 'iowa': 'IA',
        'kansas': 'KS', 'kentucky': 'KY', 'louisiana': 'LA', 'maine': 'ME', 'maryland': 'MD',
        'massachusetts': 'MA', 'michigan': 'MI', 'minnesota': 'MN', 'mississippi': 'MS', 'missouri': 'MO',
        'montana': 'MT', 'nebraska': 'NE', 'nevada': 'NV', 'new hampshire': 'NH', 'new jersey': 'NJ',
        'new mexico': 'NM', 'new york': 'NY', 'north carolina': 'NC', 'north dakota': 'ND', 'ohio': 'OH',
        'oklahoma': 'OK', 'oregon': 'OR', 'pennsylvania': 'PA', 'rhode island': 'RI', 'south carolina': 'SC',
        'south dakota': 'SD', 'tennessee': 'TN', 'texas': 'TX', 'utah': 'UT', 'vermont': 'VT',
        'virginia': 'VA', 'washington': 'WA', 'west virginia': 'WV', 'wisconsin': 'WI', 'wyoming': 'WY',
        'district of columbia': 'DC'
    }
    
    def __init__(self, cache_file, use_geocoding=True):
        self.use_geocoding = use_geocoding
        self.cache_file = Path(cache_file)
        self.cache = self.load_cache()
        if use_geocoding:
            self.geolocator = Nominatim(user_agent="pole_mapper")
        else:
            self.geolocator = None
        # Ensure all cache entries are in the correct format
        self._reformat_cache()
    
    def _reformat_cache(self):
        """Reformat all cached addresses to ensure they're in the standard format"""
        try:
            if not self.cache:
                return
                
            reformatted = False
            for key, address in list(self.cache.items()):
                # Check if address needs reformatting by looking for expected components
                if address and not all(x in address for x in [',', 'USA']):
                    parts = address.split(',')
                    # Only attempt to reformat if we have enough parts and geocoding is enabled
                    if self.use_geocoding and self.geolocator and len(parts) >= 2:
                        try:
                            lat, lon = key.split(',')
                            location = self.geolocator.reverse(f"{lat}, {lon}", timeout=10)
                            if location:
                                new_address = self.format_address(location.raw.get('address', {}))
                                if new_address:
                                    self.cache[key] = new_address
                                    reformatted = True
                        except Exception as e:
                            logging.debug(f"Failed to reformat cache entry {key}: {e}")
            
            # If any entries were reformatted, rewrite the entire cache file
            if reformatted:
                self._rewrite_cache_file()
                
        except Exception as e:
            logging.error(f"Error reformatting cache: {e}")
    
    def _rewrite_cache_file(self):
        """Rewrite the entire cache file with current cache contents"""
        try:
            with open(self.cache_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["latitude", "longitude", "address"])
                writer.writeheader()
                for key, address in self.cache.items():
                    lat, lon = key.split(',')
                    writer.writerow({
                        "latitude": lat,
                        "longitude": lon,
                        "address": address
                    })
        except Exception as e:
            logging.error(f"Error rewriting cache file: {e}")
    
    def load_cache(self):
        """Load addresses from cache file"""
        cache = {}
        if self.cache_file.exists():
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        key = f"{row['latitude']},{row['longitude']}"
                        cache[key] = row['address']
            except Exception as e:
                logging.error(f"Error loading geocoding cache: {e}")
        return cache
    
    def save_to_cache(self, key, address):
        """Save address to cache file"""
        try:
            lat, lon = key.split(',')
            # Create file with headers if it doesn't exist
            if not self.cache_file.exists():
                with open(self.cache_file, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=["latitude", "longitude", "address"])
                    writer.writeheader()
            
            # Append the new entry
            with open(self.cache_file, 'a', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["latitude", "longitude", "address"])
                writer.writerow({"latitude": lat, "longitude": lon, "address": address})
            self.cache[key] = address
        except Exception as e:
            logging.error(f"Error saving to geocoding cache: {e}")
    
    def reverse(self, lat, lon):
        """Get address from coordinates using cache or geocoding service"""
        try:
            lat, lon = float(lat), float(lon)
            key = f"{round(lat, 7)},{round(lon, 7)}"
            
            # Check cache first
            if key in self.cache:
                return self.cache[key]
            
            # If geocoding is enabled and not in cache, try geocoding service
            if self.use_geocoding and self.geolocator:
                location = self.geolocator.reverse(f"{lat}, {lon}", timeout=10)
                if location:
                    # Format address in the standard format: '903 Chittock Ave, Jackson, MI 49203, USA'
                    address = self.format_address(location.raw.get('address', {}))
                    if address:
                        # Save formatted address to cache
                        self.save_to_cache(key, address)
                        return address
            
            return ''
        except Exception as e:
            logging.error(f"Error in reverse geocoding: {e}")
            return ''
    
    def format_address(self, addr):
        """Format address in the format: '903 Chittock Ave, Jackson, MI 49203, USA'"""
        parts = []
        
        # Add house number and road
        if addr.get('house_number') and addr.get('road'):
            parts.append(f"{addr['house_number']} {addr['road']}")
        elif addr.get('road'):
            parts.append(addr['road'])
            
        # Add city
        if addr.get('city'):
            parts.append(addr['city'])
            
        # Add state and postcode
        if addr.get('state'):
            # Convert state name to abbreviation
            state_name = addr['state'].lower()
            state_code = self.STATE_CODES.get(state_name, addr['state'])
            
            state_zip = state_code
            if addr.get('postcode'):
                state_zip = f"{state_code} {addr['postcode']}"
            parts.append(state_zip)
        
        # Add country - use full 'USA' instead of 'US'
        if addr.get('country_code'):
            country = 'USA' if addr['country_code'].upper() == 'US' else addr['country_code'].upper()
            parts.append(country)
        
        return ', '.join(parts) if parts else ''