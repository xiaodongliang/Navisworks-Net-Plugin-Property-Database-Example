//------------------------------------------------------------------
// Navisworks Sample code
//------------------------------------------------------------------
//
// (C) Copyright 2011 by Autodesk Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//------------------------------------------------------------------
// Navisworks API CADOLink
// Utility class for managing the connection to an access database
//------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace NetPluginPropertyDatabaseExample
{
   class CADOLink
   {
      OleDbConnection m_connection;

      public bool connect(String dataSource) 
      {

          String conString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;; Data Source={0}", dataSource);

         m_connection = new OleDbConnection(conString);
         m_connection.Open();
         return true;
      }
      public dynamic read(String fieldName, String queryName)
      {
         // Given a known database (sheet) we can apply a named query to a specific field and return the result.
         String retString = null;
         OleDbDataReader reader;
         using (var sourceCmd = new OleDbCommand("GateHouse_LayerInfo", m_connection) { CommandType = CommandType.TableDirect })
         reader = sourceCmd.ExecuteReader();
         while (reader.Read())
         {
            int ord = reader.GetOrdinal("Name");
            if ((reader.GetString(ord)) == queryName)
            {
               object val = reader.GetValue(reader.GetOrdinal(fieldName));
               if (val != null)
               {
                  if (val.GetType() == Type.GetType("System.DateTime"))
                  {
                     DateTime dval = (DateTime)(val);
                     return dval;
                  }
                  else if (val.GetType() == Type.GetType("System.String"))
                    retString = (String)(val);
               }
               break;
            }
         } 
         return retString;
      }
      public void close()
      {
         if (m_connection != null)
            m_connection.Close();
      }
    }
}
