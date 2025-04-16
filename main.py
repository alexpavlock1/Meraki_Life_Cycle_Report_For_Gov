import os
import sys
import argparse
import asyncio
import datetime
import time
import json
import subprocess
import logging
from pptx import Presentation
import shutil

# Configure root logger to prevent debug messages from appearing in console
logging.basicConfig(level=logging.WARNING)

# Import slide modules
try:
    import clients
    CLIENTS_AVAILABLE = True
except ImportError:
    print("Warning: clients.py module not found. Dashboard summary slide will not be updated.")
    CLIENTS_AVAILABLE = False

try:
    import mx_firmware_restrictions
    MX_FIRMWARE_AVAILABLE = True
except ImportError:
    print("Warning: mx_firmware_restrictions.py module not found. MX firmware slide will not be updated.")
    MX_FIRMWARE_AVAILABLE = False

try:
    import ms_firmware_restrictions
    MS_FIRMWARE_AVAILABLE = True
except ImportError:
    print("Warning: ms_firmware_restrictions.py module not found. MS firmware slide will not be updated.")
    MS_FIRMWARE_AVAILABLE = False

try:
    import mr_firmware_restrictions
    MR_FIRMWARE_AVAILABLE = True
except ImportError:
    print("Warning: mr_firmware_restrictions.py module not found. MR firmware slide will not be updated.")
    MR_FIRMWARE_AVAILABLE = False

try:
    import mv_firmware_restrictions
    MV_FIRMWARE_AVAILABLE = True
except ImportError:
    print("Warning: mv_firmware_restrictions.py module not found. MV firmware slide will not be updated.")
    MV_FIRMWARE_AVAILABLE = False

try:
    import mg_firmware_restrictions
    MG_FIRMWARE_AVAILABLE = True
except ImportError:
    print("Warning: mg_firmware_restrictions.py module not found. MG firmware slide will not be updated.")
    MG_FIRMWARE_AVAILABLE = False

try:
    import firmware_compliance_mxmsmr
    FIRMWARE_COMPLIANCE_MXMSMR_AVAILABLE = True
except ImportError:
    print("Warning: firmware_compliance_mxmsmr.py module not found. MX/MS/MR firmware compliance slide will not be updated.")
    FIRMWARE_COMPLIANCE_MXMSMR_AVAILABLE = False

try:
    import firmware_compliance_mgmvmt
    FIRMWARE_COMPLIANCE_MGMVMT_AVAILABLE = True
except ImportError:
    print("Warning: firmware_compliance_mgmvmt.py module not found. MG/MV/MT firmware compliance slide will not be updated.")
    FIRMWARE_COMPLIANCE_MGMVMT_AVAILABLE = False

try:
    import end_of_life
    END_OF_LIFE_AVAILABLE = True
except ImportError:
    print("Warning: end_of_life.py module not found. End of Life slides will not be updated.")
    END_OF_LIFE_AVAILABLE = False

# Import the adoption module
try:
    import adoption
    ADOPTION_AVAILABLE = True
except ImportError:
    print("Warning: adoption.py module not found. Meraki Product Adoption slide will not be added.")
    ADOPTION_AVAILABLE = False

# Import the new executive summary module
try:
    import executive_summary
    EXECUTIVE_SUMMARY_AVAILABLE = True
except ImportError:
    print("Warning: executive_summary.py module not found. Executive Summary slide will not be added.")
    EXECUTIVE_SUMMARY_AVAILABLE = False
try:
    import predictive_lifecycle
    PREDICTIVE_LIFECYCLE_AVAILABLE = True
except ImportError:
    print("Warning: predictive_lifecycle.py module not found.")
    PREDICTIVE_LIFECYCLE_AVAILABLE = False
# Constants
TEMPLATE_PATH = "template.pptx"  # Path to template PPTX file, if available
OUTPUT_PATH = "meraki_report.pptx"  # Default output path

# ANSI color codes for terminal output
BLUE = '\033[94m'      # General information highlights
PURPLE = '\033[95m'    # Timer information
YELLOW = '\033[93m'    # Warnings
RED = '\033[91m'       # Errors
GREEN = '\033[92m'     # Success
RESET = '\033[0m'      # Reset to default color

def print_progress_bar(progress, total, prefix='Progress:', suffix='Complete', length=50, fill='█'):
    """
    Print a progress bar to the terminal.
    
    Args:
        progress: Current progress value
        total: Total value for 100% progress
        prefix: Text before the progress bar
        suffix: Text after the progress bar
        length: Character length of the progress bar
        fill: Character to use for the filled portion
    """
    percent = f"{100 * (progress / float(total)):.1f}%"
    filled_length = int(length * progress // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    # Clear the current line and print the progress bar
    sys.stdout.write(f"\r{BLUE}{prefix} |{GREEN}{bar}{BLUE}| {percent} {suffix}{RESET}")
    sys.stdout.flush()
    # Print a new line if progress is complete
    if progress == total:
        print()

def delete_template_slide_3(output_path):
    """
    Delete slide 3 which is just inserted with the template.
    
    Args:
        output_path: Path to the PowerPoint file
    """
    try:
        # Load presentation
        #print(f"{BLUE}Opening PowerPoint to remove slide 3...{RESET}")
        prs = Presentation(output_path)
        
        # Make sure we have enough slides
        if len(prs.slides) >= 3:
            # Get the slide 3 (index 2)
            slide = prs.slides[2]
            slide_id = slide.slide_id
            
            # Get the parent element (slide list)
            slides = prs.slides._sldIdLst
            
            # Find and remove the slide
            for i, slide_element in enumerate(slides):
                if slide_element.id == slide_id:
                    slides.remove(slide_element)
                    #print(f"{GREEN}Removed template slide 3{RESET}")
                    break
            
            # Save the updated presentation
            prs.save(output_path)
            #print(f"{GREEN}PowerPoint saved with slide 3 removed{RESET}")
        else:
            print(f"{YELLOW}Slide 3 not found in the presentation{RESET}")
    
    except Exception as e:
        print(f"{RED}Error removing slide 3: {e}{RESET}")
        import traceback
        traceback.print_exc()

def delete_slides_for_missing_devices(output_path, device_types):
    """
    Delete slides for device types that aren't present in inventory.
    
    Args:
        output_path: Path to the PowerPoint file
        device_types: Dictionary of device types with boolean values indicating presence
    """
    try:
        # Load presentation
        #print(f"{BLUE}Opening PowerPoint to remove unnecessary slides...{RESET}")
        prs = Presentation(output_path)
        
        # Map of slide indices to check and delete
        # Note: These are 0-based indices (after slide 3 is removed)
        slides_to_check = {
            2: ("MX", device_types.get('has_mx_devices', False)),
            3: ("MS", device_types.get('has_ms_devices', False)),
            4: ("MR", device_types.get('has_mr_devices', False)),
            5: ("MV", device_types.get('has_mv_devices', False)),
            6: ("MG", device_types.get('has_mg_devices', False))
        }
        
        # We need to delete from higher indices to lower to avoid index shifting
        slides_to_delete = [(idx, device_type) for idx, (device_type, exists) in slides_to_check.items() 
                            if not exists]
        
        # Sort in reverse order so we delete from the end first
        slides_to_delete.sort(reverse=True)
        
        if slides_to_delete:
            #print(f"{BLUE}Will remove {len(slides_to_delete)} slides for missing device types:{RESET}")
            for idx, device_type in slides_to_delete:
                #print(f"  - Slide {idx + 1} ({device_type} Firmware Restrictions)")
                pass
            
            # Delete slides
            for idx, device_type in slides_to_delete:
                # Check if the index is still valid
                if idx < len(prs.slides):
                    # Get the XML element for the slide
                    slide = prs.slides[idx]
                    slide_id = slide.slide_id
                    
                    # Get the parent element (slide list)
                    slides = prs.slides._sldIdLst
                    
                    # Find and remove the slide
                    for i, slide_element in enumerate(slides):
                        if slide_element.id == slide_id:
                            slides.remove(slide_element)
                            #print(f"{GREEN}Removed slide {idx + 1} ({device_type} Firmware Restrictions){RESET}")
                            break
                else:
                    print(f"{YELLOW}Slide index {idx + 1} is out of range, skipping{RESET}")
            
            # Save the updated presentation
            prs.save(output_path)
            #print(f"{GREEN}PowerPoint saved with unnecessary slides removed{RESET}")
        else:
            print(f"{BLUE}All device types present, no slides need to be removed{RESET}")
    
    except Exception as e:
        print(f"{RED}Error removing slides: {e}{RESET}")
        import traceback
        traceback.print_exc()

async def main():
    """Main orchestration function."""
    # Start timer
    start_time = time.time()
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="Generate Meraki Dashboard Report in PowerPoint")
    parser.add_argument("-o", required=True, nargs='+', 
                        help="Space-separated list of Meraki organization IDs")
    parser.add_argument("-d", "--days", type=int, default=14,
                        help="Number of days to look back for client data (1-31, default: 14)")
    parser.add_argument("--output", default=OUTPUT_PATH,
                        help=f"Output path for PowerPoint file (default: {OUTPUT_PATH})")
    parser.add_argument("--template", default=TEMPLATE_PATH,
                        help=f"Path to PowerPoint template (default: {TEMPLATE_PATH})")
    parser.add_argument("--slides", default="all",
                        help="Comma-separated list of slide numbers to generate (default: all)")
    parser.add_argument("--debug", action="store_true",
                        help="Enable verbose debug output")
    parser.add_argument("--keep-all-slides", action="store_true",
                        help="Don't remove slides for missing device types")
    parser.add_argument("--secure-connect", action="store_true",
                        help="Indicate that the organization has Secure Connect deployed")
    parser.add_argument("--umbrella", action="store_true",
                        help="Indicate that the organization has Umbrella deployed")
    parser.add_argument("--thousand-eyes", action="store_true",
                        help="Indicate that the organization has Thousand Eyes deployed")
    parser.add_argument("--spaces", action="store_true",
                        help="Indicate that the organization has Spaces deployed")
    parser.add_argument("--xdr", action="store_true",
                        help="Indicate that the organization has XDR deployed")
    parser.add_argument("--no-progress-bar", action="store_true",
                        help="Disable progress bar display")
    parser.add_argument("--export-csv", action="store_true",
                        help="Export firmware compliance data to CSV files")
    
    args = parser.parse_args()
    
    # Enable debug mode if requested
    debug_mode = args.debug
    
    # Setup progress tracking
    use_progress_bar = not args.no_progress_bar
    progress_current = 0
    progress_total = 0  # Will be calculated based on slides to generate
    
    # Validate days input
    if args.days < 1 or args.days > 31:
        print("Error: Days must be between 1 and 31. Using default (14).")
        args.days = 14
    
    # Update paths based on arguments
    output_path = args.output
    template_path = args.template
    
    # Print slide numbers for clarity
    # print(f"{YELLOW}Slide reference for clarity:{RESET}")
    # print(f"Slide 1: Title slide (updated with organization names)")
    # print(f"Slide 2: Dashboard summary (updated by clients.py)")
    # print(f"Slide 3: MX firmware restrictions (updated by mx_firmware_restrictions.py)")
    # print(f"Slide 4: MS firmware restrictions (updated by ms_firmware_restrictions.py)")
    # print(f"Slide 5: MR firmware restrictions (updated by mr_firmware_restrictions.py)")
    # print(f"Slide 6: MV firmware restrictions (updated by mv_firmware_restrictions.py)")
    # print(f"Slide 7: MG firmware restrictions (updated by mg_firmware_restrictions.py)")
    # print(f"Slide 8: Firmware Compliance MX/MS/MR (updated by firmware_compliance_mxmsmr.py)")
    # print(f"Slide 9: Firmware Compliance MG/MV/MT (updated by firmware_compliance_mgmvmt.py)")
    # print(f"Slide 10: End of Life Products (updated by end_of_life.py)")
    # print(f"Slide 11: Device Models and EOL Dates (updated by end_of_life.py)")
    # 
    # if ADOPTION_AVAILABLE:
    #     print(f"Additional: Meraki Product Adoption slide (dynamically added by adoption.py)")
    #     
    # if EXECUTIVE_SUMMARY_AVAILABLE:
    #     print(f"Additional: Executive Summary slide (added after all other slides, then moved to position 2)")
    
    # Determine which slides to generate
    available_slides = []
    if CLIENTS_AVAILABLE:
        available_slides.append(2)
    if MX_FIRMWARE_AVAILABLE:
        available_slides.append(3)
    if MS_FIRMWARE_AVAILABLE:
        available_slides.append(4)
    if MR_FIRMWARE_AVAILABLE:
        available_slides.append(5)
    if MV_FIRMWARE_AVAILABLE:
        available_slides.append(6)
    if MG_FIRMWARE_AVAILABLE:
        available_slides.append(7)
    if FIRMWARE_COMPLIANCE_MXMSMR_AVAILABLE:
        available_slides.append(8)
    if FIRMWARE_COMPLIANCE_MGMVMT_AVAILABLE:
        available_slides.append(9)
    if END_OF_LIFE_AVAILABLE:
        available_slides.append(10)
        available_slides.append(11)
        
    # Create a list for additional special slides
    special_slides = []
    if ADOPTION_AVAILABLE:
        special_slides.append('product_adoption')
    if EXECUTIVE_SUMMARY_AVAILABLE:
        special_slides.append('executive_summary')
    if PREDICTIVE_LIFECYCLE_AVAILABLE:  # Added this line - always include predictive lifecycle
        special_slides.append('predictive_lifecycle')
        
    if args.slides.lower() == 'all':
        slides_to_generate = available_slides + special_slides
    else:
        try:
            requested_slides = []
            for slide in args.slides.split(','):
                if slide.lower() in ['product_adoption', 'executive_summary']:
                    requested_slides.append(slide.lower())
                else:
                    requested_slides.append(int(slide))
            slides_to_generate = requested_slides
        except ValueError:
            print(f"{RED}Error: Invalid slide numbers. Using all available slides.{RESET}")
            slides_to_generate = available_slides + special_slides
    
    if not slides_to_generate:
        #print(f"{RED}No valid slides specified or no slide modules available. Exiting.{RESET}")
        return

    print(f"\n{BLUE}Starting Meraki Dashboard Report Generation{RESET}")
    
    # Set up progress tracking
    if use_progress_bar:
        # Set up a fixed total of 100 steps for easier percentage tracking
        progress_total = 100
        
        # Initialize progress
        progress_current = 0
        print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    # print(f"Organizations: {args.o}")
    # print(f"Days for client data: {args.days}")
    # print(f"Template: {template_path}")
    # print(f"Output: {output_path}")
    # print(f"Slides to generate: {slides_to_generate}")
    
    # Variables to store all collected data
    dashboard_stats = None
    all_inventory_devices = []
    org_names = {}  # Store organization names
    
    # Variables to track device types
    device_types = {
        'has_mx_devices': False,
        'has_ms_devices': False,
        'has_mr_devices': False,
        'has_mv_devices': False,
        'has_mg_devices': False
    }
    
    # PHASE 1: Data Collection
    # First collect data for all slides without updating PowerPoint
    if 2 in slides_to_generate and CLIENTS_AVAILABLE:
        data_start_time = time.time()
        print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Collecting dashboard data...{RESET}")
        
        try:
            # Import functions from clients.py
            from clients import get_inventory_devices, get_api_key, AdaptiveRateLimiter
            from clients import get_networks, get_dashboard_stats, filter_incompatible_networks
            from clients import get_client_stats, get_organization_names
            import meraki.aio
            
            # Get API key and create rate limiter
            api_key = get_api_key()
            rate_limiter = AdaptiveRateLimiter()
            
            # Set up Meraki client
            async with meraki.aio.AsyncDashboardAPI(
                api_key=api_key,
                suppress_logging=True,
                maximum_retries=3,
                base_url="https://api.gov-meraki.com/api/v1"
            ) as aiomeraki:
                # Get organization names
                print(f"{BLUE}Getting organization names...{RESET}")
                org_names = await get_organization_names(aiomeraki, args.o, rate_limiter)
                # print(f"Organization names: {org_names}")
                
                # Get networks
                print(f"{BLUE}Getting networks...{RESET}")
                all_networks = []
                for org_id in args.o:
                    try:
                        networks = await get_networks(aiomeraki, org_id, rate_limiter)
                        all_networks.extend(networks)
                        # print(f"Found {len(networks)} networks in organization {org_id}")
                    except Exception as e:
                        print(f"{RED}Error retrieving networks for org {org_id}: {e}{RESET}")
                
                # Get network IDs
                network_ids = [network['id'] for network in all_networks]
                
                # Filter incompatible networks
                #print(f"{BLUE}Filtering networks...{RESET}")
                valid_network_ids = await filter_incompatible_networks(network_ids, all_networks)
                
                # Get dashboard statistics
                print(f"{BLUE}Getting dashboard statistics...{RESET}")
                dash_stats = await get_dashboard_stats(aiomeraki, args.o, rate_limiter)
                
                # Get client statistics
                print(f"{BLUE}Getting client statistics...This may take some time. Please be patient.{RESET}")
                
                # Update progress bar before getting client statistics
                if use_progress_bar:
                    # Set progress to 4%
                    progress_current = 4
                    print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
                client_stats = await get_client_stats(aiomeraki, valid_network_ids, rate_limiter, args.days)
                
                # Combine all stats
                dashboard_stats = {**dash_stats, **client_stats}
                
                # Get inventory devices for firmware restriction slides
                print(f"{BLUE}Getting inventory devices...{RESET}")
                
                # Update progress bar before getting inventory devices
                if use_progress_bar:
                    # Set progress to 8%
                    progress_current = 8
                    print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
                for org_id in args.o:
                    try:
                        devices = await get_inventory_devices(aiomeraki, org_id, rate_limiter)
                        all_inventory_devices.extend(devices)
                        # print(f"{GREEN}Retrieved {len(devices)} inventory devices from org {org_id}{RESET}")
                    except Exception as e:
                        print(f"{RED}Error retrieving inventory for org {org_id}: {e}{RESET}")
                
                # print(f"{GREEN}Total inventory devices retrieved: {len(all_inventory_devices)}{RESET}")
                
                # Detect which device types are present in inventory
                device_types['has_mx_devices'] = any(device.get('model', '').upper().startswith('MX') for device in all_inventory_devices)
                device_types['has_ms_devices'] = any(device.get('model', '').upper().startswith('MS') for device in all_inventory_devices)
                device_types['has_mr_devices'] = any(device.get('model', '').upper().startswith('MR') or device.get('model', '').upper().startswith('CW') for device in all_inventory_devices)
                device_types['has_mv_devices'] = any(device.get('model', '').upper().startswith('MV') for device in all_inventory_devices)
                device_types['has_mg_devices'] = any(device.get('model', '').upper().startswith('MG') for device in all_inventory_devices)
                
                # Print device type summary
                # print(f"\n{BLUE}Device Types Summary:{RESET}")
                # print(f"MX Security Appliances: {'Present' if device_types['has_mx_devices'] else 'Not Present'}")
                # print(f"MS Switches: {'Present' if device_types['has_ms_devices'] else 'Not Present'}")
                # print(f"MR Access Points: {'Present' if device_types['has_mr_devices'] else 'Not Present'}")
                # print(f"MV Cameras: {'Present' if device_types['has_mv_devices'] else 'Not Present'}")
                # print(f"MG Cellular Gateways: {'Present' if device_types['has_mg_devices'] else 'Not Present'}")
        
        except Exception as e:
            print(f"{RED}Error collecting data: {e}{RESET}")
            import traceback
            traceback.print_exc()
        
        data_time = time.time() - data_start_time
        print(f"{PURPLE}Data collection completed in {data_time:.2f} seconds{RESET}")
        
        # Update progress bar
        if use_progress_bar:
            progress_current = 12.5
            print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    # PHASE 2: PowerPoint Updates
    # Now update PowerPoint with the collected data
    
    # First, copy template to output if needed
    if template_path != output_path and not os.path.exists(output_path):
        try:
            import shutil
            shutil.copy2(template_path, output_path)
            #print(f"{GREEN}Created output file from template{RESET}")
        except Exception as e:
            print(f"{RED}Error copying template to output: {e}{RESET}")
    
    # Delete template slide 3 before adding content
    delete_template_slide_3(output_path)
    
    # Update progress bar
    if use_progress_bar:
        progress_current = 15
        print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    # Update slide 2 (Dashboard Summary) using the update_clients.py script
    if dashboard_stats and 2 in slides_to_generate and CLIENTS_AVAILABLE:
        slide_start_time = time.time()
        print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 2...{RESET}")
        
        try:
            # Convert dashboard_stats to JSON
            stats_json = json.dumps(dashboard_stats)
            
            # Also convert org_names to JSON for the update script
            org_names_json = json.dumps(org_names)
            
            # Call the update_slides.py script with the stats, days, and org_names
            result = subprocess.run(
                ["python3", "update_clients.py", template_path, output_path, stats_json, str(args.days), org_names_json],
                capture_output=True,
                text=True
            )
            
            # Print output from the script
            print(result.stdout)

            if result.stderr:
                print(f"{RED}Errors from update script:{RESET}")
                print(result.stderr)

            if result.returncode == 0:
                #print(f"{GREEN}Updated dashboard summary in PowerPoint{RESET}")
                pass
            else:
                print(f"{RED}Error updating dashboard summary{RESET}")
                
        except Exception as e:
            print(f"{RED}Error updating slide 2: {e}{RESET}")
            import traceback
            traceback.print_exc()
        
        slide_time = time.time() - slide_start_time
        print(f"{PURPLE}Slide 2 update completed in {slide_time:.2f} seconds{RESET}")
        
        # Update progress bar
        if use_progress_bar:
            progress_current = 18.8
            print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    # Update slide 3 (MX Firmware Restrictions) using mx_firmware_restrictions.py
    if all_inventory_devices and 3 in slides_to_generate and MX_FIRMWARE_AVAILABLE:
        if device_types['has_mx_devices']:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 3...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call mx_firmware_restrictions's generate function
                if hasattr(mx_firmware_restrictions, 'generate'):
                    await mx_firmware_restrictions.generate(
                        api_client,
                        output_path,  # Use potentially already updated file as template
                        output_path,
                        inventory_devices=all_inventory_devices
                    )
                    print(f"{GREEN}Updated MX firmware restrictions in PowerPoint{RESET}")
                else:
                    print(f"{RED}mx_firmware_restrictions.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 3: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 3 update completed in {slide_time:.2f} seconds{RESET}")
            
            # Update progress bar
            if use_progress_bar:
                # After slide 3, progress to 25%
                progress_current = 25
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
        else:
            print(f"\n{YELLOW}Skipping slide 3 - No MX devices found in inventory{RESET}")
            
            # Update progress bar even if we skip
            if use_progress_bar and 3 in slides_to_generate:
                # After slide 3 (even if skipped), progress to 25%
                progress_current = 25
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    # Update slide 4 (MS Firmware Restrictions) using ms_firmware_restrictions.py
    if all_inventory_devices and 4 in slides_to_generate and MS_FIRMWARE_AVAILABLE:
        if device_types['has_ms_devices']:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 4...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call ms_firmware_restrictions's generate function
                if hasattr(ms_firmware_restrictions, 'generate'):
                    await ms_firmware_restrictions.generate(
                        api_client,
                        output_path,  # Use potentially already updated file as template
                        output_path,
                        inventory_devices=all_inventory_devices
                    )
                    print(f"{GREEN}Updated MS firmware restrictions in PowerPoint{RESET}")
                else:
                    print(f"{RED}ms_firmware_restrictions.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 4: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 4 update completed in {slide_time:.2f} seconds{RESET}")
            
            if use_progress_bar:
                progress_current = 31.2
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
        else:
            print(f"\n{YELLOW}Skipping slide 4 - No MS devices found in inventory{RESET}")
            
            if use_progress_bar and 4 in slides_to_generate:
                progress_current = 31.2
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    # Update slide 5 (MR Firmware Restrictions) using mr_firmware_restrictions.py
    if all_inventory_devices and 5 in slides_to_generate and MR_FIRMWARE_AVAILABLE:
        if device_types['has_mr_devices']:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 5...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call mr_firmware_restrictions's generate function
                if hasattr(mr_firmware_restrictions, 'generate'):
                    await mr_firmware_restrictions.generate(
                        api_client,
                        output_path,  # Use potentially already updated file as template
                        output_path,
                        inventory_devices=all_inventory_devices
                    )
                    print(f"{GREEN}Updated MR firmware restrictions in PowerPoint{RESET}")
                else:
                    print(f"{RED}mr_firmware_restrictions.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 5: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 5 update completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping slide 5 - No MR devices found in inventory{RESET}")
    
    # Update slide 6 (MV Firmware Restrictions) using mv_firmware_restrictions.py
    if all_inventory_devices and 6 in slides_to_generate and MV_FIRMWARE_AVAILABLE:
        if device_types['has_mv_devices']:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 6...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call mv_firmware_restrictions's generate function
                if hasattr(mv_firmware_restrictions, 'generate'):
                    await mv_firmware_restrictions.generate(
                        api_client,
                        output_path,
                        output_path,
                        inventory_devices=all_inventory_devices
                    )
                    print(f"{GREEN}Updated MV firmware restrictions in PowerPoint{RESET}")
                else:
                    print(f"{RED}mv_firmware_restrictions.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 6: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 6 update completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping slide 6 - No MV devices found in inventory{RESET}")
    
    # Update slide 7 (MG Firmware Restrictions) using mg_firmware_restrictions.py
    if all_inventory_devices and 7 in slides_to_generate and MG_FIRMWARE_AVAILABLE:
        if device_types['has_mg_devices']:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 7...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call mg_firmware_restrictions's generate function
                if hasattr(mg_firmware_restrictions, 'generate'):
                    await mg_firmware_restrictions.generate(
                        api_client,
                        output_path,  # Use potentially already updated file as template
                        output_path,
                        inventory_devices=all_inventory_devices
                    )
                    print(f"{GREEN}Updated MG firmware restrictions in PowerPoint{RESET}")
                else:
                    print(f"{RED}mg_firmware_restrictions.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 7: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 7 update completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping slide 7 - No MG devices found in inventory{RESET}")
    
    # Update slide 8 (Firmware Compliance MX/MS/MR) using firmware_compliance_mxmsmr.py
    if 8 in slides_to_generate and FIRMWARE_COMPLIANCE_MXMSMR_AVAILABLE:
        # This slide requires networks data, which should be available from clients.py
        if 'all_networks' in locals() and all_networks:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 8...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                if hasattr(firmware_compliance_mxmsmr, 'generate'):
                    await firmware_compliance_mxmsmr.generate(
                        api_client,
                        output_path,
                        output_path,
                        networks=all_networks,
                        export_csv=args.export_csv
                    )
                    print(f"{GREEN}Updated Firmware Compliance MX/MS/MR slide in PowerPoint{RESET}")
                else:
                    print(f"{RED}firmware_compliance_mxmsmr.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 8: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 8 update completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping slide 8 - No networks data available{RESET}")
    
    # Update slide 9 (Firmware Compliance MG/MV/MT) using firmware_compliance_mgmvmt.py
    if 9 in slides_to_generate and FIRMWARE_COMPLIANCE_MGMVMT_AVAILABLE:
        # This slide requires networks data, which should be available from clients.py
        if 'all_networks' in locals() and all_networks:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 9...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                if hasattr(firmware_compliance_mgmvmt, 'generate'):
                    await firmware_compliance_mgmvmt.generate(
                        api_client,
                        output_path,
                        output_path,
                        networks=all_networks,
                        export_csv=args.export_csv
                    )
                    print(f"{GREEN}Updated Firmware Compliance MG/MV/MT slide in PowerPoint{RESET}")
                else:
                    print(f"{RED}firmware_compliance_mgmvmt.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 9: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 9 update completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping slide 9 - No networks data available{RESET}")
    
    # Update slide 10 (End of Life Products) using end_of_life.py
    if 10 in slides_to_generate and END_OF_LIFE_AVAILABLE:
        # This slide requires inventory devices data, which should be available
        if all_inventory_devices:
            slide_start_time = time.time()
                # Silenced to reduce terminal output
            # print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating slide 10...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call end_of_life's generate function
                if hasattr(end_of_life, 'generate'):
                    await end_of_life.generate(
                        api_client,
                        output_path,
                        output_path,
                        inventory_devices=all_inventory_devices,
                        networks=all_networks if 'all_networks' in locals() else None
                    )
                    print(f"{GREEN}Updated End of Life Products slide in PowerPoint{RESET}")
                else:
                    print(f"{RED}end_of_life.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error updating slide 10: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 10 update completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping slide 10 - No inventory device data available{RESET}")

    # Update slide 11 (Device Models and EOL Dates) using end_of_life.py
    if 11 in slides_to_generate and END_OF_LIFE_AVAILABLE:
        # This slide requires inventory devices data, which should be available
        if all_inventory_devices:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Creating slide 11...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call end_of_life's generate_detail_slide function
                if hasattr(end_of_life, 'generate_detail_slide'):
                    await end_of_life.generate_detail_slide(
                        api_client,
                        output_path,  # Use potentially already updated file as template
                        output_path,
                        inventory_devices=all_inventory_devices,
                        networks=all_networks if 'all_networks' in locals() else None
                    )
                    print(f"{GREEN}Created Device Models and EOL Dates slide in PowerPoint{RESET}")
                else:
                    print(f"{RED}end_of_life.py doesn't have generate_detail_slide function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error creating slide 11: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Slide 11 creation completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping slide 11 - No inventory device data available{RESET}")
    
    # Create Meraki Product Adoption slide using adoption.py
    products_adoption_data = None
    if 'product_adoption' in slides_to_generate and ADOPTION_AVAILABLE:
        # This slide requires inventory devices data, which should be available
        if all_inventory_devices:
            slide_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Creating Meraki Product Adoption slide...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Create manual configuration dictionary - this can be loaded from a config file or args
                manual_config = {
                    'Secure Connect': False,
                    'Umbrella Secure Internet Gateway': False,
                    'Thousand Eyes': False, 
                    'Spaces': False,
                    'XDR': False
                }
                
                # Check for manual override arguments
                if hasattr(args, 'secure_connect'):
                    manual_config['Secure Connect'] = args.secure_connect
                if hasattr(args, 'umbrella'):
                    manual_config['Umbrella Secure Internet Gateway'] = args.umbrella
                if hasattr(args, 'thousand_eyes'):
                    manual_config['Thousand Eyes'] = args.thousand_eyes
                if hasattr(args, 'spaces'):
                    manual_config['Spaces'] = args.spaces
                if hasattr(args, 'xdr'):
                    manual_config['XDR'] = args.xdr
                
                # Store products adoption data for potential use in executive summary
                products_adoption_data = {
                    'MX': device_types['has_mx_devices'],
                    'MS': device_types['has_ms_devices'],
                    'MR': device_types['has_mr_devices'],
                    'MG': device_types['has_mg_devices'],
                    'MV': device_types['has_mv_devices'],
                    'MT': any(device.get('model', '').upper().startswith('MT') for device in all_inventory_devices),
                    'Secure Connect': manual_config['Secure Connect'],
                    'Umbrella Secure Internet Gateway': manual_config['Umbrella Secure Internet Gateway'],
                    'Thousand Eyes': manual_config['Thousand Eyes'],
                    'Spaces': manual_config['Spaces'],
                    'XDR': manual_config['XDR']
                }
                
                # Call adoption's generate function
                if hasattr(adoption, 'generate'):
                    await adoption.generate(
                        api_client,
                        output_path,
                        output_path,
                        inventory_devices=all_inventory_devices,
                        networks=all_networks if 'all_networks' in locals() else None,
                        manual_config=manual_config
                    )
                    print(f"{GREEN}Created Meraki Product Adoption slide in PowerPoint{RESET}")
                else:
                    print(f"{RED}adoption.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error creating Meraki Product Adoption slide: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            slide_time = time.time() - slide_start_time
            print(f"{PURPLE}Meraki Product Adoption slide creation completed in {slide_time:.2f} seconds{RESET}")
        else:
            print(f"\n{YELLOW}Skipping Meraki Product Adoption slide - No inventory device data available{RESET}")
    
    # After all slides have been updated, delete unnecessary slides
    if all_inventory_devices and not args.keep_all_slides:
        delete_start_time = time.time()
        print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Removing slides for missing device types...{RESET}")
        
        # Delete slides for missing device types
        delete_slides_for_missing_devices(output_path, device_types)
        
        delete_time = time.time() - delete_start_time
        print(f"{PURPLE}Slide deletion completed in {delete_time:.2f} seconds{RESET}")
        
        if use_progress_bar:
            progress_current = 37.5
            print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    # Update title slide with organization names if not done through update_slides.py
    if org_names and not (2 in slides_to_generate and CLIENTS_AVAILABLE):
        title_start_time = time.time()
        print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Updating title slide...{RESET}")
        
        title_time = time.time() - title_start_time
        print(f"{PURPLE}Title slide update completed in {title_time:.2f} seconds{RESET}")
            
    # Finally, after ALL other slides are generated, create the Executive Summary slide
    firmware_compliance_data = None
    eol_data = None
    
    # Try to extract firmware compliance data from slides 8-9
    if FIRMWARE_COMPLIANCE_MXMSMR_AVAILABLE:
        try:
            firmware_compliance_data = {
                'MX': {'Good': 0, 'Warning': 0, 'Critical': 0, 'Total': 0, 'latest': None},
                'MS': {'Good': 0, 'Warning': 0, 'Critical': 0, 'Total': 0, 'latest': None},
                'MR': {'Good': 0, 'Warning': 0, 'Critical': 0, 'Total': 0, 'latest': None}
            }
            
        except Exception as e:
            print(f"{YELLOW}Could not extract firmware compliance data: {e}{RESET}")
    
    # Try to extract EOL data from slides 10-11
    if END_OF_LIFE_AVAILABLE:
        try:
            # Get the actual EOL data from the documentation
            from end_of_life import get_eol_info_from_doc
            eol_data, last_updated, is_from_doc = get_eol_info_from_doc()
            

            # print(f"{GREEN}Using EOL data from documentation for predictive lifecycle{RESET}")
            # if "MX100" in eol_data:
            #     print(f"{GREEN}MX100 EOL data for predictive lifecycle: {eol_data['MX100']}{RESET}")
        except Exception as e:
            print(f"{YELLOW}Could not fetch EOL data: {e}, using fallback{RESET}")
            # Only use fallback data if fetch fails
            try:
                from end_of_life import EOL_FALLBACK_DATA
                eol_data = EOL_FALLBACK_DATA
            except Exception as e2:
                print(f"{YELLOW}Could not import EOL data: {e2}, using None{RESET}")
    
    if 'executive_summary' in slides_to_generate and EXECUTIVE_SUMMARY_AVAILABLE:
        if all_inventory_devices:
            exec_summary_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Creating Executive Summary slide...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Call the executive summary's generate function
                if hasattr(executive_summary, 'generate'):
                    await executive_summary.generate(
                        api_client,
                        output_path,
                        output_path,
                        inventory_devices=all_inventory_devices,
                        networks=all_networks if 'all_networks' in locals() else None,
                        dashboard_stats=dashboard_stats,
                        firmware_stats=firmware_compliance_data,
                        eol_data=eol_data,
                        products=products_adoption_data
                    )
                    print(f"{GREEN}Created Executive Summary slide in PowerPoint (inserted at position 2){RESET}")
                else:
                    print(f"{RED}executive_summary.py doesn't have generate function{RESET}")
            
            except Exception as e:
                print(f"{RED}Error creating Executive Summary slide: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            exec_summary_time = time.time() - exec_summary_start_time
            print(f"{PURPLE}Executive Summary slide creation completed in {exec_summary_time:.2f} seconds{RESET}")
            
            # Update progress bar - Executive Summary complete
            if use_progress_bar:
                # After Executive Summary, progress to 43.8%
                progress_current = 43.8
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
        else:
            print(f"\n{YELLOW}Skipping Executive Summary slide - No inventory device data available{RESET}")
            
            # Update progress bar even if we skip
            if use_progress_bar and 'executive_summary' in slides_to_generate:
                # After Executive Summary (even if skipped), progress to 43.8%
                progress_current = 43.8
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    if 'predictive_lifecycle' in slides_to_generate and PREDICTIVE_LIFECYCLE_AVAILABLE:
        # This slide uses inventory devices, EOL data, and networks
        if all_inventory_devices:
            lifecycle_start_time = time.time()
            print(f"\n{PURPLE}[{time.strftime('%H:%M:%S')}] Creating Predictive Lifecycle Management slides...{RESET}")
            
            try:
                # Create simple API client
                class SimpleApiClient:
                    def __init__(self, org_ids):
                        self.org_ids = org_ids
                        self.dashboard = None
                
                api_client = SimpleApiClient(args.o)
                
                # Extract EOL data if available
                eol_data = None
                if END_OF_LIFE_AVAILABLE:
                    try:
                        # Get the actual EOL data from the documentation
                        from end_of_life import get_eol_info_from_doc
                        eol_data, last_updated, is_from_doc = get_eol_info_from_doc()
                        
                        # print(f"{GREEN}Using EOL data from documentation for predictive lifecycle{RESET}")
                        # if "MX100" in eol_data:
                        #     print(f"{GREEN}MX100 EOL data for predictive lifecycle: {eol_data['MX100']}{RESET}")
                        pass
                    except Exception as e:
                        print(f"{YELLOW}Could not fetch EOL data: {e}, using fallback{RESET}")
                        # Only use fallback data if fetch fails
                        try:
                            from end_of_life import EOL_FALLBACK_DATA
                            eol_data = EOL_FALLBACK_DATA
                        except Exception as e2:
                            print(f"{YELLOW}Could not import EOL data: {e2}, using None{RESET}")
                                
                if use_progress_bar:
                    progress_current = 66
                    print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
                
                # Generate the predictive lifecycle slides
                # Note: not specifying a position, so they'll be added at the end
                await predictive_lifecycle.generate(
                    api_client,
                    output_path,
                    output_path,
                    inventory_devices=all_inventory_devices,
                    networks=all_networks if 'all_networks' in locals() else None,
                    eol_data=eol_data
                )
                
                if use_progress_bar:
                    progress_current = 75
                    print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
                
                print(f"{GREEN}Added Predictive Lifecycle Management slides to the end of the presentation{RESET}")
                
            except Exception as e:
                print(f"{RED}Error creating Predictive Lifecycle Management slides: {e}{RESET}")
                import traceback
                traceback.print_exc()
            
            lifecycle_time = time.time() - lifecycle_start_time
            print(f"{PURPLE}Predictive Lifecycle slides creation completed in {lifecycle_time:.2f} seconds{RESET}")
            

            if use_progress_bar:
                progress_current = 95
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
        else:
            print(f"\n{YELLOW}Skipping Predictive Lifecycle slides - No inventory device data available{RESET}")
            

            if use_progress_bar and 'predictive_lifecycle' in slides_to_generate:
                progress_current = 95
                print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    # Calculate total script execution time
    total_time = time.time() - start_time
    
    if use_progress_bar and progress_current < 100:
        progress_current = 100
        print_progress_bar(progress_current, progress_total, prefix='Overall Progress:', suffix='Complete')
    
    print(f"\n{PURPLE}Total script execution time: {total_time:.2f} seconds{RESET}")
    print(f"\n{BLUE}Dashboard Report created successfully at {output_path}{RESET}")

def run_individual_slide(slide_num):
    """Helper function to run a single slide generator for debugging."""
    if slide_num == 2 and CLIENTS_AVAILABLE:
        print(f"{YELLOW}Running clients.py directly for debugging{RESET}")
        # Create a simple test case
        test_orgs = ["123456"]
        asyncio.run(clients.main_async(test_orgs, 7, TEMPLATE_PATH, OUTPUT_PATH))
    elif slide_num == 3 and MX_FIRMWARE_AVAILABLE:
        print(f"{YELLOW}Running mx_firmware_restrictions.py directly for debugging{RESET}")
        # Create a simple test case
        test_orgs = ["123456"]
        asyncio.run(mx_firmware_restrictions.main_async(test_orgs, TEMPLATE_PATH, OUTPUT_PATH))

    elif slide_num == "adoption" and ADOPTION_AVAILABLE:
        print(f"{YELLOW}Running adoption.py directly for debugging{RESET}")
        # Create a simple test case
        test_orgs = ["123456"]
        test_manual_config = {
            'Secure Connect': True,
            'Umbrella Secure Internet Gateway': False,
            'Thousand Eyes': True,
            'Spaces': False,
            'XDR': False
        }
        asyncio.run(adoption.main_async(test_orgs, TEMPLATE_PATH, OUTPUT_PATH, test_manual_config))
    elif slide_num == "executive_summary" and EXECUTIVE_SUMMARY_AVAILABLE:
        print(f"{YELLOW}Running executive_summary.py directly for debugging{RESET}")
        test_orgs = ["123456"]
        asyncio.run(executive_summary.main_async(test_orgs, TEMPLATE_PATH, OUTPUT_PATH))
    else:
        print(f"{RED}Invalid slide number or module not available{RESET}")

if __name__ == "__main__":
    # Check for special debug flags
    if len(sys.argv) > 1 and sys.argv[1] == "--debug-clients":
        run_individual_slide(2)
    elif len(sys.argv) > 1 and sys.argv[1] == "--debug-mx":
        run_individual_slide(3)
    # [Other debug checks...]
    elif len(sys.argv) > 1 and sys.argv[1] == "--debug-adoption":
        run_individual_slide("adoption")
    elif len(sys.argv) > 1 and sys.argv[1] == "--debug-executive-summary":
        run_individual_slide("executive_summary")
    else:
        asyncio.run(main())
