# Webtoon/Comic Update Checker
--------------------------------
A python 2.7 script to check updates on manga, comics, webtoons and opens if available. 

Currently only compatible with [webtoons.com](http://www.webtoons.com/)

__Usage:__
```
python komicupdate.py
```
---
__Required dependencies:__
* 'index.xml' file in directory
* lxml (3.8.0)
* requests (2.18.4)
* xmltodict (0.11.0) 

index.xml should follow this format:
```
<?xml version="1.0"?>
<list>
	<comic name="<name>" title="<title_no>" genre="<genre>" language="<language_id>" type="<type_id>">
		<link url="<base_url>" episode="<ep_no>" chapter="<episode_no>"/>
	</comic>
	...
</list>

```
where,
```
<!-- Webtoon Link-->
<base_url>/<language_id>/<genre>/<name>/ep-<ep_no>/viewer?title_no=<title_no>&episode_no=<episode_no>
```

---

###__Timeline:__

_Last Updated:_ Oct 23, 2017

10/21/17 - Initial creation, currently only compatible for webtoon.com
10/23/17 - Re-implemented with xmltodict

-------
####__TO DO:__
- Check for empty 'input.xml' prior to comic iteration
- Create a logger class
- Separate code into other files
- Allow for multiple for multiple sources, see comic_type

