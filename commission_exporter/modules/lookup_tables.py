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