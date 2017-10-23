#!/usr/bin/python

#==========================================================================
# Webtoon/Comic Update Checker
# by Jessa Mae Alcantara
# 
# Purpose: Script to check updates on manga, comics,
# webtoons and opens if available.
#
# Last Updated: Oct 23, 2017
# -------------------------------------------------------------------------
# 10/21/17 - Initial creation, currently only compatible for webtoon.com
# 10/23/17 - Re-implemented with xmltodict
#
# TO DO:
# - Check for empty 'input.xml' prior to comic iteration
# - Create a logger class
# - Separate code into other files
# - Allow for multiple 
#==========================================================================

import requests
import xmltodict
import webbrowser
from lxml import etree
    
# BUILD FUNCTIONS
# -----------------------------

# builds url for webtoon service
def build_url(comic):
    name = format_name(comic['@name'])
    url = comic['@url'] + '/' + comic['@language'] + '/' + comic['@genre'] + '/' + name + '/ep-' + comic['@episode'] + '/viewer?title_no=' + comic['@title'] + '&episode_no=' + comic['@chapter']
    return url

# build webtoon link xml
def build_link(link):
    link_xml = '<link url="http://www.webtoons.com" episode="' + link['@episode'] + '" chapter="' + link['@chapter'] + '"/>'
    return link_xml

# build individual comic xml
def build_comic(comic):
    link = build_link(comic['link'])
    comic_xml = '<comic name="'+ comic['@name'] + '" title="' + comic['@title'] + '" genre="' + comic['@genre'] + '" language="' + comic['@language'] + '" type="'+ comic['@type'] + '">\n\t\t'+ link + '\n\t</comic>\n'
    return comic_xml


# build document with updated current info
def build_xml(comics):
    xml_header = '<?xml version="1.0"?>\n<list>\n\t'
    xml_footer = '</list>\n'
    xml_content = ''
    
    for comic in comics:
        individual_comic = build_comic(comic)
        xml_content += individual_comic
        
    xml = xml_header + xml_content + xml_footer
    return xml

# HELPER FUNCTIONS:
# -----------------------------

# format comic name to proper url format
def format_name(name):
    formatted_name = name.replace(' ', '-').lower()
    formatted_name.lower()
    return formatted_name

# format comic name to proper xml format
def format_name(name):
    formatted_name = name.replace('-', ' ').lower()
    return formatted_name


# format webtoon dict for url creation/log
def webt_format(format_dict):
    link_dict = format_dict['link']
    del format_dict['link']
    format_dict.update(link_dict)
    return format_dict

# APP FUNCTIONS:
# -----------------------------

# parses xml document to comics dictionary
def comicsParse():
    with open('input.xml') as xml_doc:
        xml = xmltodict.parse(xml_doc.read())

    comics = xml['list']['comic']
    return comics

# get values for next chapter
def next_chap(link):
    episode = str(int(link['@episode']) + 1)
    chapter = str(int(link['@chapter']) + 1)
    link['@episode'] = episode
    link['@chapter'] = chapter
    return link

# check http response for url
def check_new(url):
    req = requests.get(url)
    status = req.status_code
    if (status == 200):
        return True
    else:
        return False

# log for updated chapter
def log_update(comic):
    # Need to account for index offset in url
    ch = str(int(comic['@chapter']) + 1)
    print('Chapter '+ ch + ' - '+ comic['@name'][0].upper() + comic['@name'][1:] + ' has been updated.')

# log for unchanged chapter
def log_unchanged(comic):
    # Need to account for index offset in url
    print(comic['@name'][0].upper() + comic['@name'][1:] + ' is still waiting on the next chapter.')
          
# application function
def main():
    urls = []
    print( 'Opening your \'input.xml\' file...')
    comics = comicsParse()

    # iterate through ea comic checking status and updating/opening comic if new chp is available
    print( 'Processing your specified comics...')
    for comic in comics:
        comic_type = comic['@type']

        # store next chp info in temp as copy
        tmp = comic.copy()
        current_link = tmp['link'].copy()
        new_link = next_chap(current_link)
        tmp['link'] = new_link

        # build url to next chp info
        format_tmp = webt_format(tmp) 
        url = build_url(format_tmp)

        # check if chp is available
        status = check_new(url)
        if (status):
            log_update(format_tmp)
            comic['link'] = new_link
            urls.append(url)
        else:
            log_unchanged(comic)
                
    # open all new chapters
    for new_ch in urls:
        webbrowser.open(new_ch)
                
    # update 'input.xml'
    print( 'Updating your \'input.xml\' file...')
    xml = build_xml(comics)
    with open('input.xml','wb') as doc:
        doc.write(xml)


main()
