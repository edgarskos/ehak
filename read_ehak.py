#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
import os, time

import xlrd

sys.path.append("/home/rk/py/pywikibot/core")
import pywikibot

import logging


def from_this_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

def pageExists(wPage):
    try:
        text = wPage.get()
        return True
    except pywikibot.NoPage: # First except, prevent empty pages
        return False
    except pywikibot.IsRedirectPage: # second except, prevent redirect
        pywikibot.output(u'%s is a redirect!' % wPage.title() )
        logging.warning('[[%s]] is a redirect!' % (wPage.title() ))
        return False
    except pywikibot.Error: # third exception, take the problem and print
        pywikibot.output(u"Some error, skipping..")
        return False 
        
def getWikiArticle(site, nimi, vald, tyyp, tyypNimi):
    
    #print tyyp
    if tyyp == 1:
        valdCatName = u'Kategooria:%s' % nimi
    else:
        valdCatName = u'Kategooria:%s vald' % vald
    #print valdCatName
    catVald = pywikibot.Page(site, valdCatName)
    
    if tyyp == 0:      #maakond
        logging.info('%s is %s. Skipping.' % (nimi, tyypNimi) )
        return False
    elif tyyp == 1:    #vald
        tplVald = pywikibot.Page(site, u'Mall:EestiVald')
        
        pageName = nimi
        wPage = pywikibot.Page(site, pageName)
        if pageExists(wPage):
            if tplVald in wPage.templates() and catVald in wPage.categories():
                return wPage
        
        logging.warning('No wiki page found for: %s.' % (nimi) )
        return False
    elif tyyp == 6:    #linnaosa
        logging.info('%s is %s. Skipping.' % (nimi, tyypNimi) )
        return False
    else:              #asula
        tplAsula = pywikibot.Page(site, u'Mall:EestiAsula')
        
        if len(vald) > 1:
            pageName = u'%s (%s)' % (nimi, vald)
            wPage = pywikibot.Page(site, pageName)
            if pageExists(wPage):
                if tplAsula in wPage.templates() and catVald in wPage.categories():
                    return wPage
                    
        pageName = nimi
        wPage = pywikibot.Page(site, pageName)
        if pageExists(wPage):
            if tplAsula in wPage.templates() and catVald in wPage.categories():
                return wPage

        pageName = u'%s %s' % (nimi, tyypNimi)
        wPage = pywikibot.Page(site, pageName)
        if pageExists(wPage):
            if tplAsula in wPage.templates() and catVald in wPage.categories():
                return wPage
                      
        logging.warning('No wiki page found for: %s.' % (nimi) )
        print 'No wiki page found for: %s.' % (nimi)
        return False
        
    logging.warning('Got to the end of getWikiArticle() function!' )
    return False

def getDataPageTitle(wPageTitle):
    site = pywikibot.Site("et", "wikipedia")
    wPage = pywikibot.Page(site, wPageTitle)
    dataPage = pywikibot.ItemPage.fromPage(wPage)
    if dataPage.exists():
        return dataPage.title()


def editDataPage(site, wPage, kood, nimi, tyyp, tyypNimi, vald, maakond):
    
    source_url = 'http://metaweb.stat.ee/get_classificator_file.htm?id=3765296&siteLanguage=et'
    
    repo = site.data_repository()
    
    claims_rules = {
        'P17': 'Q191', # Country = Estonia
    }

    tyypPages = {
        3: 'Q3374262', 
        7: 'Q3744870', 
        8: 'Q532',
        4: 'Q3957', 
        1: 'Q15284', #vald 
        5: 'Q15715391', #vallasisene linn
    }
    
    dataPage = pywikibot.ItemPage.fromPage(wPage)
    settlement = tyypPages[tyyp]
    
    assert settlement
    claims_rules['P31'] = settlement
    
    if tyyp == 1:
        isInAdminUnit = maakond
    else:
        isInAdminUnit = '%s vald' % vald
        
    isInAdminDPTitle = getDataPageTitle(isInAdminUnit)
    
    assert isInAdminDPTitle
    claims_rules['P131'] = isInAdminDPTitle
    
    assert kood
    claims_rules['P1140'] = kood
    
    if dataPage.exists():
        dataPage.get()
        
        #print dataPage.descriptions
        if 'et' in dataPage.descriptions:
            dummy = 1
            #print dataPage.descriptions['et']
        else:
            etDesc = u'%s' % (tyypNimi)
            if len(vald) > 2:
                etDesc = u'%s %s vallas' % (etDesc, vald)
            mkDesc = maakond.replace(u' maakond', u'')
            etDesc = u'%s %smaal' % (etDesc, mkDesc)
            mydescriptions = {u'et': etDesc}
            pywikibot.output('Adding desc: %s' % (etDesc))
            dataPage.editDescriptions(mydescriptions)

        
        for pid in claims_rules:
            if pid in dataPage.claims: #check if existing claim matches
                if issubclass(type(dataPage.claims[pid][0].getTarget()), pywikibot.ItemPage):
                    WDvalue = dataPage.claims[pid][0].getTarget().title()
                else:
                    WDvalue = dataPage.claims[pid][0].getTarget()
                    
                if WDvalue != claims_rules[pid]:
                    if pid == 'P131':    #asub haldusüksuses
                        claim = dataPage.claims[pid][0]
                        target = pywikibot.ItemPage(repo, claims_rules[pid])
                        pywikibot.output('Overwriting %s --> %s' % (claim.getID(), claim.getTarget()))
                        claim.changeTarget(target)
                        #getting edit conflict error otherwise
                        time.sleep(5)
                        #url as source
                        source = pywikibot.Claim(repo, 'P854')
                        source.setTarget(source_url)
                        claim.addSource(source)
                    else:
                        print ('[[%s]] property %s value doesnt match with WD value' % (wPage.title(), pid ))
                        logging.warning('[[%s]] property %s value doesnt match with WD value' % (wPage.title(), pid ))
            else:                      #add new claim
                print claims_rules[pid]
                claim = pywikibot.Claim(repo, pid)
                
                
                if claims_rules[pid].startswith( 'Q' ):
                    target = pywikibot.ItemPage(repo, claims_rules[pid])
                else:
                    target = claims_rules[pid]
                claim.setTarget(target)
                pywikibot.output('Adding %s --> %s' % (claim.getID(), claim.getTarget()))
                dataPage.addClaim(claim)
                #url as source
                source = pywikibot.Claim(repo, 'P854')
                source.setTarget(source_url)
                claim.addSource(source)


                
    else:
        print 'ERROR: NO DATA PAGE'
        logging.warning('[[%s]]: no data page in Wikidata' % (wPage.title() ))


###  main

logging.basicConfig(filename='ehak.log',level=logging.DEBUG)
logging.basicConfig(format='%(asctime)s %(message)s')
  
book = xlrd.open_workbook(from_this_dir('EHAK2014v1.xls'), formatting_info=True)
sheet = book.sheet_by_name('EHAK2014v1')

site = pywikibot.Site("et", "wikipedia")
 
print sheet.nrows
NROWS = 150

for rowNum in range(56, NROWS):
    
    kood = sheet.cell_value(rowNum, 0)
    nimi = sheet.cell_value(rowNum, 1)
    nimi = nimi.replace(u' küla', u'')
    nimi = nimi.replace(u' alevik', u'')
    nimi = nimi.replace(u' alev', u'')
    nimi = nimi.replace(u' vallasisene linn', u'')
    nimi = nimi.replace(u' linn', u'')
    tyyp = sheet.cell_value(rowNum, 3)
    tyyp = int(tyyp)
    tyypNimi = sheet.cell_value(rowNum, 4)
    vald = sheet.cell_value(rowNum, 6)
    vald = vald.replace(u' vald', u'')
    maakond = sheet.cell_value(rowNum, 8)
    print '%s, %s, %s, %s, %s, %s' % (kood, nimi, tyyp, tyypNimi, vald, maakond)
    
    wPage = getWikiArticle(site, nimi, vald, tyyp, tyypNimi)
    if wPage:
        print 'wiki artikkel: %s' % (wPage.title())
        editDataPage(site, wPage, kood, nimi, tyyp, tyypNimi, vald, maakond)
        #time.sleep(10)
        
