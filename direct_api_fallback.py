"""
Direct API Fallback Module for Meraki Client Fetching

This module provides a fallback mechanism for fetching Meraki client data when the SDK approach fails. This is most common with API timeouts seen when trying to run with SDK.
It maintains persistent state across runs to remember problematic networks.
"""

import os
import asyncio
import time
import json
import logging
import aiohttp
from datetime import datetime, timedelta

# Set up logging
logger = logging.getLogger('direct_api_fallback')
logger.setLevel(logging.DEBUG)

# Create logs directory if it doesn't exist
if not os.path.exists('logs'):
    os.makedirs('logs')

# Create a file handler
file_handler = logging.FileHandler('logs/direct_api_fallback.log')
file_handler.setLevel(logging.DEBUG)

# Create a formatter
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# Add the handler to the logger
logger.addHandler(file_handler)

# Prevent log messages from propagating to the console
logger.propagate = False

# In-memory cache for the current run
_problematic_networks = set()

# Persistent state file
STATE_FILE = '.meraki_problematic_networks.json'

def _load_problematic_networks():
    """Load the list of problematic networks from a persistent JSON file."""
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, 'r') as f:
                networks = json.load(f)
                logger.debug(f"Loaded {len(networks)} problematic networks from state file")
                return set(networks)
    except Exception as e:
        logger.error(f"Error loading problematic networks: {e}")
    
    return set()

def _save_problematic_networks():
    """Save the current list of problematic networks to a JSON file."""
    try:
        with open(STATE_FILE, 'w') as f:
            json.dump(list(_problematic_networks), f)
            logger.debug(f"Saved {len(_problematic_networks)} problematic networks to state file")
    except Exception as e:
        logger.error(f"Error saving problematic networks: {e}")

# Load previously identified problematic networks on module import
_problematic_networks = _load_problematic_networks()

def mark_as_problematic(network_id):
    """Mark a network as problematic for future runs."""
    global _problematic_networks
    if network_id not in _problematic_networks:
        _problematic_networks.add(network_id)
        _save_problematic_networks()
        logger.info(f"Marked network {network_id} as problematic")
    return True

def is_problematic(network_id):
    """Check if a network is known to be problematic."""
    return network_id in _problematic_networks

async def _direct_api_call(network_id, t0, t1, api_key):
    """Make a direct API call to get clients, with detailed error handling."""
    url = f"https://api.gov-meraki.com/api/v1/networks/{network_id}/clients"
    
    params = {
        "t0": t0,
        "t1": t1,
        "perPage": 5000  # Maximum allowed
    }
    
    headers = {
        "X-Cisco-Meraki-API-Key": api_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "Python Meraki Direct API Client"
    }
    
    logger.debug(f"DIRECT API REQUEST: {url} with t0={t0}, t1={t1}")
    
    try:
        async with aiohttp.ClientSession() as session:
            try:
                start_time = time.time()
                async with session.get(url, params=params, headers=headers, timeout=60) as response:
                    response_time = time.time() - start_time
                    logger.debug(f"DIRECT API RESPONSE: Status {response.status} in {response_time:.2f}s")
                    
                    if response.status == 200:
                        data = await response.json()
                        logger.debug(f"DIRECT API SUCCESS: Got {len(data)} clients")
                        return data
                    elif response.status == 429:
                        error_text = await response.text()
                        logger.error(f"DIRECT API RATE LIMIT: {error_text[:200]}")
                        # Add a 1-second delay before raising the exception
                        await asyncio.sleep(1)
                        raise Exception(f"API rate limit: {response.status}, {error_text[:200]}")
                    else:
                        error_text = await response.text()
                        logger.error(f"DIRECT API ERROR: Status {response.status}, {error_text[:200]}...")
                        raise Exception(f"API error: {response.status}, {error_text[:200]}")
            except asyncio.TimeoutError as e:
                logger.error(f"DIRECT API TIMEOUT after {time.time() - start_time:.2f}s")
                raise e
            except asyncio.CancelledError:
                # Explicitly catch and re-raise CancelledError
                logger.error(f"DIRECT API CANCELLED for network {network_id}")
                raise
            except Exception as e:
                logger.error(f"DIRECT API EXCEPTION: {type(e).__name__} - {str(e)}")
                raise e
    except asyncio.CancelledError:
        # Explicitly catch and re-raise CancelledError at the session level
        logger.error(f"DIRECT API SESSION CANCELLED for network {network_id}")
        raise
    except Exception as e:
        logger.error(f"DIRECT API SESSION EXCEPTION: {type(e).__name__} - {str(e)}")
        raise e

async def _chunked_api_call(network_id, t0_str, t1_str, api_key, chunk_hours=1, max_retries=3):
    """Use time chunking to get clients when a direct call fails."""
    logger.debug(f"Starting chunked API call with {chunk_hours} hour chunks")
    
    # Parse start and end times
    t0 = datetime.strptime(t0_str, "%Y-%m-%dT%H:%M:%SZ")
    t1 = datetime.strptime(t1_str, "%Y-%m-%dT%H:%M:%SZ")
    
    all_clients = []
    current_time = t0
    
    while current_time < t1:
        # Calculate next chunk end time
        next_time = min(current_time + timedelta(hours=chunk_hours), t1)
        
        # Format times for API
        chunk_t0 = current_time.strftime("%Y-%m-%dT%H:%M:%SZ")
        chunk_t1 = next_time.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        # Implement retries with exponential backoff
        retry_count = 0
        success = False
        
        while retry_count < max_retries and not success:
            try:
                # Add a small random delay between chunks to avoid rate limiting
                await asyncio.sleep(0.3)
                
                # Make the API call for this chunk
                logger.debug(f"Calling API for chunk {chunk_t0} to {chunk_t1} (retry {retry_count}/{max_retries})")
                chunk_clients = await _direct_api_call(network_id, chunk_t0, chunk_t1, api_key)
                
                # Add clients from this chunk to the overall list
                all_clients.extend(chunk_clients)
                logger.debug(f"Got {len(chunk_clients)} clients for chunk {chunk_t0} to {chunk_t1}")
                
                # Success - break out of retry loop
                success = True
                
            except asyncio.CancelledError:
                # Handle cancellation by propagating it up
                logger.error(f"Chunk API call cancelled for network {network_id}")
                raise
            except Exception as e:
                logger.error(f"Error in chunk {chunk_t0} to {chunk_t1} (retry {retry_count}/{max_retries}): {e}")
                retry_count += 1
                
                if retry_count >= max_retries:
                    # If we've reached max retries and still using larger chunks, try with smaller chunks
                    if chunk_hours > 0.25:  # Don't go smaller than 15 minutes
                        logger.debug(f"Retrying with smaller chunks ({chunk_hours/2} hours)")
                        try:
                            # Try with smaller chunks for this specific time range
                            smaller_chunk_clients = await _chunked_api_call(
                                network_id,
                                chunk_t0,
                                chunk_t1,
                                api_key,
                                chunk_hours=chunk_hours/2,
                                max_retries=max_retries
                            )
                            all_clients.extend(smaller_chunk_clients)
                            success = True
                        except asyncio.CancelledError:
                            # If the smaller chunks are cancelled, propagate it up
                            raise
                        except Exception as e2:
                            logger.error(f"Smaller chunks also failed: {e2}")
                            # Continue to next time chunk
                    else:
                        # If already at minimum chunk size, skip this chunk
                        logger.warning(f"Skipping chunk {chunk_t0} to {chunk_t1} after all retries failed")
                else:
                    # Exponential backoff between retries
                    backoff_time = 0.5 * (2 ** retry_count)
                    logger.debug(f"Backing off for {backoff_time:.2f}s before retry {retry_count+1}")
                    await asyncio.sleep(backoff_time)
        
        # Move to the next chunk
        current_time = next_time
    
    return all_clients

async def fallback_handler(network_id, t0, t1, api_key):
    """Main handler for direct API fallback.
    
    Tries three approaches in sequence:
    1. Direct API call for the full time range
    2. If #1 fails, try chunked API calls with 1-hour chunks
    3. If a chunk fails, recursively try smaller chunks down to 15 minutes
    """
    logger.info(f"Fallback handler called for network {network_id}, t0={t0}, t1={t1}")
    
    try:
        # First try a direct API call for the full time range
        logger.debug("Attempting direct API call for full time range")
        clients = await _direct_api_call(network_id, t0, t1, api_key)
        logger.info(f"Direct API fallback succeeded for network {network_id} - got {len(clients)} clients")
        return clients
        
    except asyncio.CancelledError:
        # Handle cancellation explicitly and return an empty list instead of propagating
        logger.error(f"Direct API call cancelled for network {network_id}, returning empty list")
        return []
    except Exception as e:
        logger.warning(f"Direct API call failed: {e}, falling back to chunked approach")
        
        # If direct call fails, try chunked approach
        try:
            clients = await _chunked_api_call(network_id, t0, t1, api_key)
            logger.info(f"Chunked API fallback succeeded for network {network_id} - got {len(clients)} clients")
            return clients
        except asyncio.CancelledError:
            # Handle cancellation explicitly and return an empty list
            logger.error(f"Chunked API call cancelled for network {network_id}, returning empty list")
            return []
        except Exception as e2:
            logger.error(f"All API fallback methods failed for network {network_id[-4:]}: {e2}")
            # Return empty list as fallback
            return []