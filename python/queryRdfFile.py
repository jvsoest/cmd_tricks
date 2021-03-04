import argparse
import rdflib
import os

# Parse arguments
parser = argparse.ArgumentParser(description='Perform a SPARQL query on an RDF file')
parser.add_argument('rdfFilePath', help='Path to the RDF file')
parser.add_argument('sparqlQuery', help='Path to the SPARQL query file, Query template, or the actual query string')
inputArgs = parser.parse_args()

# Create graph
g = rdflib.Graph()

# Read TTL file
result = g.parse(inputArgs.rdfFilePath, format=rdflib.util.guess_format(inputArgs.rdfFilePath))

# Determine whether to:
#  a) read a sparql query file
#  b) read a sample sparql query
#  c) parse the actual string as query
query = None
queriesFolder = os.path.join(os.path.expanduser('~'), "Documents", "Repositories", "cmd_tricks", "python", "sparqlQueries")
if query is None:
    if os.path.exists(inputArgs.sparqlQuery):
        print("Using given query file")
        with open(inputArgs.sparqlQuery) as f:
            query = f.read()
if query is None:
    fileToSearch = os.path.join(queriesFolder, inputArgs.sparqlQuery) + ".sparql"
    print(fileToSearch)
    if os.path.exists(fileToSearch):
        print("Using existing query")
        with open(fileToSearch) as f:
            query = f.read()
if query is None:
    print("Using given string as query")
    query = inputArgs.sparqlQuery

# Execute the actual query
qResult = g.query(query)

# Loop over query results, and visualize them
for row in qResult.bindings:
    print(row)