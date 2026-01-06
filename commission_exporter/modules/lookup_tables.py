from typing import Dict, List, Set, Tuple, Optional, Any, Iterable

def get_doc_check_criteria():
    checks = {
        'pb': 'Before Photo',
        'pa': 'After Photo',
        'pr': 'Receipt Photo',
        'qd': 'Quote Description',
        'qs': 'Quote Signed',
        'qe': 'Quote Emailed',
        'id': 'Invoice Description',
        'is': 'Invoice Signed',
        'ins': 'Invoice Not Signed (Client Offsite)',
        'ie': 'Invoice Emailed',
        '5s': '5 Star Review',
    }
    return checks

def get_tenants():
    tenants = {
        'FoxtrotWhiskey (NSW)': 'foxtrotwhiskey',
        'BravoGolf (QLD)': 'bravogolf',
        'MikeEcho (VIC)': 'mikeecho',
        'SierraDelta (WA)': 'sierradelta',
        'AlphaBravo (Old NSW)': 'alphabravo',
        'EchoZulu (Old QLD)': 'echozulu',
        'VictorTango (Old VIC)': 'victortango',
    }
    return tenants

def get_state_from_tenant(tenant: Optional[str] = None):
    tenants = {
        'foxtrotwhiskey': 'NSW',
        'bravogolf': 'QLD',
        'mikeecho': 'VIC',
        'sierradelta': 'WA',
        'alphabravo': 'NSW',
        'echozulu': 'QLD',
        'victortango': 'VIC',
    }
    if tenant: return tenants[tenant]
    return tenants

def get_tenant_from_state(state: Optional[str] = None) -> dict[str,list] | list:
    mapping = {
        'NSW': ['foxtrotwhiskey', 'alphabravo'],
        'QLD': ['bravogolf', 'echozulu'],
        'VIC': ['mikeecho', 'victortango'],
        'WA': ['sierradelta'],
    }
    if state: return mapping[state]
    return mapping

def get_all_payment_types():
    output = {
                'AMEX': 'CC',
                'Applied Payment for AR': 'DK',
                'Cash': 'Cash',
                'Check': 'DK',
                'Credit Card': 'CC',
                'EFT/Bank Transfer': 'EFT',
                'Humm - Finance Fee': 'PP',
                'Humm Payment Plan': 'PP',
                'Zip Payment Plan': 'PP',
                'Zip - Finance Fee': 'PP',
                'Imported Default Credit Card': 'DK',
                'MasterCard': 'CC',
                'Visa': 'CC',
                'Payment Plan': 'PP',
                'Payment Plan - Fee': 'PP',
                'Processed in ServiceM8': 'DK',
                'Refund (check)': 'REF',
                'Refund (credit card)': 'REF',
                '': 'DK'
            }
    return output