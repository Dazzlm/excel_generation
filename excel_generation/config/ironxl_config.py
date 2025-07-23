# excel_generation/config/ironxl_config.py

from ironxl import License # <-- Import the License class
import os

def configurar_ironxl():
    """
    Configures the IronXL license key.
    """
    key = os.environ.get("LicenseKeyIronXL", "YOUR-LICENSE-KEY") # Use your actual key
    
    # Correct way to set the license
    License.set_LicenseKey(key)
    
    print("IronXL license configured successfully.")