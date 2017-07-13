#!/usr/bin/python
#
# Converts a taxonomy in YAML to a Excel file: first sheet lists the vocabularies, 
# following sheets list the terms for each of the vocabularies.
#
# Requires: 
# * XlsxWriter https://xlsxwriter.readthedocs.io/
# * PyYaml http://pyyaml.org/
#

import argparse
import yaml
import xlsxwriter

parser = argparse.ArgumentParser()
parser.add_argument('--taxos', '-t', nargs='+', required=True, help='Taxonomy YML file paths.')
parser.add_argument('--out', '-o', required=True, help='Taxonmy Excel file path.')
parser.add_argument('--verbose', '-v', action='store_true')
args = parser.parse_args()

class Taxonomy(object):
	def __init__(self, values):
		self.name = values["name"]
		self.author = values["author"]
		self.license = values["license"]
		self.title = values["title"]
		self.description = values["description"]
		self.vocabularies = values["vocabularies"]

def taxonomy_constructor(loader, node):
	return Taxonomy(loader.construct_mapping(node, deep=True))

yaml.add_constructor('tag:yaml.org,2002:org.obiba.opal.core.domain.taxonomy.Taxonomy', taxonomy_constructor)

def write_taxonomy(taxonomy, wss, rows=[1, 1, 1]):
	txws = wss[0]
	vcws = wss[1]
	trws = wss[2]
	txrow = rows[0]
	vcrow = rows[1]
	trrow = rows[2]
	#print [txrow, vcrow, trrow]
	# write taxo
	if args.verbose:
		print taxonomy.name
	txws.write(txrow, 0, taxonomy.name)
	txws.write(txrow, 1, taxonomy.title.get('en'))
	txws.write(txrow, 2, taxonomy.title.get('fr'))
	txws.write(txrow, 3, taxonomy.description.get('en'))
	txws.write(txrow, 4, taxonomy.description.get('fr'))
	txws.write(txrow, 5, taxonomy.author)
	txws.write(txrow, 6, taxonomy.license)
	txrow = txrow + 1
	# write vocabularies
	for vocabulary in taxonomy.vocabularies:
		if args.verbose:
			print "  " + vocabulary.get('name')
		vcws.write(vcrow, 0, taxonomy.name)
		vcws.write(vcrow, 1, vocabulary.get('name'))
		vcws.write(vcrow, 2, vocabulary.get('title', {}).get('en'))
		vcws.write(vcrow, 3, vocabulary.get('title', {}).get('fr'))
		vcws.write(vcrow, 4, vocabulary.get('description', {}).get('en'))
		vcws.write(vcrow, 5, vocabulary.get('description', {}).get('fr'))
		vcws.write(vcrow, 6, vocabulary.get('repeatable', '0'))
		vcrow = vcrow + 1
		# write terms
		for term in vocabulary.get('terms', []):
			if args.verbose:
				print "    " + term.get('name')
			trws.write(trrow, 0, taxonomy.name)
			trws.write(trrow, 1, vocabulary.get('name'))
			trws.write(trrow, 2, term.get('name'))
			trws.write(trrow, 3, term.get('title', {}).get('en'))
			trws.write(trrow, 4, term.get('title', {}).get('fr'))
			trws.write(trrow, 5, term.get('description', {}).get('en'))
			trws.write(trrow, 6, term.get('description', {}).get('fr'))
			trws.write(trrow, 7, term.get('keywords', {}).get('en'))
			trws.write(trrow, 8, term.get('keywords', {}).get('fr'))
			trrow = trrow + 1
	#print [txrow, vcrow, trrow]
	return [txrow, vcrow, trrow]

# prepare workbook
workbook = xlsxwriter.Workbook(args.out)
txws = workbook.add_worksheet('Taxonomies')
txws.write('A1', 'name')
txws.write('B1', 'title:en')
txws.write('C1', 'title:fr')
txws.write('D1', 'description:en')
txws.write('E1', 'description:fr')
txws.write('F1', 'author')
txws.write('G1', 'license')
vcws = workbook.add_worksheet('Vocabularies')
vcws.write('A1', 'taxonomy')
vcws.write('B1', 'name')
vcws.write('C1', 'title:en')
vcws.write('D1', 'title:fr')
vcws.write('E1', 'description:en')
vcws.write('F1', 'description:fr')
vcws.write('G1', 'repeatable')
trws = workbook.add_worksheet('Terms')
trws.write('A1', 'taxonomy')
trws.write('B1', 'vocabulary')
trws.write('C1', 'name')
trws.write('D1', 'title:en')
trws.write('E1', 'title:fr')
trws.write('F1', 'description:en')
trws.write('G1', 'description:fr')
trws.write('H1', 'keywords:en')
trws.write('I1', 'keywords:fr')

# read taxonomies
rows=[1,1,1]
for taxo in args.taxos:
	with open(taxo, 'r') as stream:
	    try:
	    	taxonomy = yaml.load(stream)
	        rows = write_taxonomy(taxonomy, wss=[txws, vcws, trws], rows=rows)
	    except yaml.YAMLError as exc:
	        print(exc)

# close workbook
workbook.close()