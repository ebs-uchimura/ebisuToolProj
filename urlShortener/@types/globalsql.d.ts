/**
 * globalsql.d.ts
 **
 * function：グローバル宣言
**/

export { };

declare global {
  // count arguments
  interface countargs {
    table: string;
    columns: string[];
    values: any[][];
    fields?: string[];
  }

  // select arguments
  interface selectargs {
    table: string;
    columns: string[];
    values: any[][];
    limit?: number;
    offset?: number;
    spancol?: string;
    span?: number;
    order?: string;
    reverse?: boolean;
    fields?: string[];
  }

  // update arguments
  interface updateargs {
    table: string;
    setcol: string[];
    setval: any[];
    selcol: string[];
    selval: any[];
    spancol?: string;
    spanval?: number;
  }

  // insert arguments
  interface insertargs {
    table: string;
    columns: string[];
    values: any[];
  }
}