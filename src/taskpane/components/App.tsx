import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { buildClientSchema, getIntrospectionQuery } from "graphql";
import Header from './Header';
import Progress from './Progress';
import GraphiQLExplorer from "graphiql-explorer";
import { makeDefaultArg, getDefaultScalarArgValue } from "./CustomArgs";
import ApolloClient from 'apollo-boost';
import gql from 'graphql-tag';
import "./App.css";

const uri = 'https://api.graph.cool/simple/v1/swapi';

const client = new ApolloClient({uri});

// const getSpaceConcat = (query) => query.join(' ');

const getObjectValuesArray = (obj,keys) => keys.map(e => obj[e]);

const getObjKeyArray2level = (obj) => {
  var keys = Object.keys(obj).map(k => Object.keys(obj[k][0])).reduce((acc, val) => acc.concat(val), []);
  return keys.filter(e => e !== '__typename');
}

const getObjFirstKey = (obj) => Object.keys(obj)[0]

const createTable = async (query) => {
  // const schema = await introspectSchema(link);

  try {
    await Excel.run(async context => {
      var response = {};
      try {
        response = await client.query({query: gql`
        ${query}
        `});
      } catch (e) {
        console.log(e);
      }
      const data = response['data'];
      const keys = getObjKeyArray2level(data);

      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var range = currentWorksheet.getRange("A1");
      range = range.getResizedRange(0, keys.length-1);
      var table = currentWorksheet.tables.add(range, true /*hasHeaders*/);
      table.name = "ExpensesTable";

      table.getHeaderRowRange().values =
          [keys];

      table.rows.add(null /*add at the end*/, 
        data[getObjFirstKey(data)].map(p => getObjectValuesArray(p,keys)));

      range.values = [[]];
      await context.sync();

    });
  } catch (error) {
    console.error(error);
  }
}

const reset = async () => {
  // const schema = await introspectSchema(link);

  try {
    await Excel.run(async context => {

      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var table = currentWorksheet.tables.getItemAt(0);
      table.delete();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}



function fetcher(params: Object) {
  return fetch(
    uri,
    {
      method: "POST",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify(params)
    }
  )
    .then(function(response) {
      return response.text();
    })
    .then(function(responseBody) {
      try {
        return JSON.parse(responseBody);
      } catch (e) {
        return responseBody;
      }
    });
}
const DEFAULT_QUERY = `# shift-option/alt-click on a query below to jump to it in the explorer
# option/alt-click on a field in the explorer to select all subfields
query MyQuery {
}
`;

export default function App(props) {
  const {
    title,
    isOfficeInitialized,
  } = props;
  const [schema, setSchema] = React.useState(null);
  const [query, setQuery] = React.useState(DEFAULT_QUERY);
  const [explorerIsOpen, setExplorerIsOpen] = React.useState(true);


  React.useEffect(() => {
    fetcher({
      query: getIntrospectionQuery()
    }).then(result => {
      setSchema(buildClientSchema(result.data));
    });
  });

  const _handleEditQuery = (query: string): void => {
    setQuery(query);
    console.log(query);
  }

  const _handleToggleExplorer = () => {
    setExplorerIsOpen(!explorerIsOpen);
  };

  const click = () => {
    createTable(query);
  }
  const clear = () => {
    reset();
  }

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo='assets/logo-filled.png'
          message='Please sideload your addin to see app body.'
        />
      );
    }

    return (
      <div className='ms-welcome'>
        <Header logo='assets/icon.png' title={props.title} message='StarWars GraphQL API' />
        <div className="graphiql-container">
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={click}>Query</Button><br/>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={clear}>Delete</Button>
        <GraphiQLExplorer
          schema={schema}
          query={query}
          onEdit={_handleEditQuery}
          onRunOperation={_ => {}
          }
          explorerIsOpen={explorerIsOpen}
          onToggleExplorer={_handleToggleExplorer}
          getDefaultScalarArgValue={getDefaultScalarArgValue}
          makeDefaultArg={makeDefaultArg}
        />

      </div>
        </div>
    );
  }
// }
