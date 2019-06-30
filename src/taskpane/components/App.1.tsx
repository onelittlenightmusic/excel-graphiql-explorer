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

const getSpaceConcat = (query) => query.join(' ');

const getObjectValuesArray = (obj,query) => query.map(e => obj[e]);


const createTable = async (query) => {
  // const schema = await introspectSchema(link);

  try {
    await Excel.run(async context => {

      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var range = currentWorksheet.getRange("A1");
      range = range.getResizedRange(0, query.length-1);
      var table = currentWorksheet.tables.add(range, true /*hasHeaders*/);
      table.name = "ExpensesTable";

      table.getHeaderRowRange().values =
          [query];

      var data = await client.query({
            query: gql`
              {
                allPersons {
                  ${getSpaceConcat(query)}
                }
              }
            `,
          });
      table.rows.add(null /*add at the end*/, 
        data.data.allPersons.map(p => getObjectValuesArray(p,query)));
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
  const [gender, setGender] = React.useState(true);
  const [id, setId] = React.useState(true);
  const [name, setName] = React.useState(true);
  const [mass, setMass] = React.useState(true);
  const [schema, setSchema] = React.useState(null);
  const [query, setQuery] = React.useState(DEFAULT_QUERY);
  const [explorerIsOpen, setExplorerIsOpen] = React.useState(true);

  const getQueryArray = () => {
    const state = {id, name, mass, gender};
    return Object.keys(state).filter(k => state[k]);
    // return ["id","name","mass","gender"].filter(k => this[k]);
  };

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
    createTable(getQueryArray());
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
        <form>
        <input
            name="checkId"
            type="checkbox"
            checked={id} 
            onChange={()=>setId(!id)}/>
              id<br/>
        <input
            name="checkName"
            type="checkbox"
            checked={name} 
            onChange={()=>setName(!name)}/>
              Name<br/>
              <input
            name="checkMass"
            type="checkbox"
            checked={mass} 
            onChange={()=>setMass(!mass)}/>
              Mass<br/>
              <input
            name="checkColar"
            type="checkbox"
            checked={gender} 
            onChange={()=>setGender(!gender)}/>
              Gender<br/>

        </form>
        </div>
    );
  }
// }
