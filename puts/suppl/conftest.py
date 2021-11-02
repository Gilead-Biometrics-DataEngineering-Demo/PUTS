# -*- coding: utf-8 -*-
import math
import re
import os
import pytest
from puts import puts


@pytest.fixture()
def fix_test(request):
    files = request.param
    # Setup each test
    wrkdir = os.getcwd()  
    for f in files:
        os.system("rm -f -- " + f)
    yield 
    # Teardown each test
    for f in files:
       os.system("rm -f -- " + f)
    os.chdir(wrkdir)  
 
    
@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    # Get result and user requirement of each test
    outcome = yield
    result = outcome.get_result()

    if result.when == 'call':
        setattr(item, "urid", re.sub(r'[^\d]','',item._obj.__doc__.split('\n')[0].strip()) )
        setattr(item, "user_requirement", item._obj.__doc__.split('\n')[1].strip() )
        if result.passed:
            setattr(item, "Pass_fail", 'PASS' )
        else:
            setattr(item, "Pass_fail", 'FAIL' )
       
            
def pytest_sessionfinish(session, exitstatus):
    # Display additional information after all tests
    print()
    print()
    print(' {:6}   {:100}   {:6} '.format('urid', 
                                            'user_requirement', 
                                            'result'))
    for item in session.items:
        # format each line of the report
        widths = [6, 100, 6]
        datarow = [item.urid, 
                    item.user_requirement,
                    item.Pass_fail + 'ED']
        maxl = max([math.ceil(len(dat)/wid) for dat, wid in zip(datarow, widths)]) # number of lines needed to split up long strings
        for x in range(1,maxl+1):
            print(' '.join(' {:{width}} '.format(str(dat).ljust(maxl)[(x-1)*wid:x*wid], 
                                                        width=wid) for dat, wid in zip(datarow, widths)))
    print() 
    
    # Generate docx and pdf forms   
    puts.pdf_frm02060(session.items)
    puts.docx_frm02060(session.items)