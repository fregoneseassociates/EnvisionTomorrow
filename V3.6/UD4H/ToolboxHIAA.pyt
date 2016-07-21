"""
################################################################################
# UD4H Health Impact Assessment Application (HIAA) Toolbox
# Get EPAT Baseline Data Tool
# Author: Thomas York (sustainablegrowth at gmail dot com)
# Date: May 26 2016
# Version: 1.0
#
License:
Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this project except in compliance with the License. You may obtain a copy of the
License at

   http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed
under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR
CONDITIONS OF ANY KIND, either express or implied. See the License for the
specific language governing permissions and limitations under the License.
################################################################################
"""
# http://desktop.arcgis.com/en/arcmap/10.3/analyze/creating-tools/a-template-for-python-toolboxes.htm
# http://desktop.arcgis.com/en/arcmap/10.3/analyze/creating-tools/a-quick-tour-of-python-toolboxes.htm
import archook
archook.get_arcpy_sp()
import arcpy
import os
import json
import urllib2
import urllib

_hmv = "EPAP"
_clientid = "info@frego.com"
#_geoidfn = "GEOID10"
_hiaa_url_default = "http://api.ud4htools.com/hmapi_post_custom_inputs/" + _hmv + "/"
_hiaa_url_detail = "hmapi_post_custom_inputs"
_hiaa_url_summary = "hmapi_get_summary_json"
#_pre_post_json_file = "hiaa_post.json"

# excluding only int4 'notrdata',
_epap_fields = { 'EPAT':
['totpop2010','tothhs2010','p_wrkage','pct_autoo0','totemp2010','totwrk2010',
'arealnd_ac','areaunp_ac','popdens_ac','empdens_ac','empentropy','trpequilib','ntwkdenped',
'intrsndens','empbytrans','jobacc45tr','retailempl','forst_nlcd','natrl_nlcd','treec_nlcd',
'opens_nlcd','pct_childr','pct_adults','pct_senior','pct_popmal','pct_popfem','pct_nohsed',
'pct_hseduc','pct_2ycoll','pct_4ypcol','pct_worker','pct_notwrk','pct_hhkids','pct_nokids',
'pct_ownocc','pct_rentoc','avg_hhsize','pct_lowinc','pct_medinc','pct_higinc','topenspace'],
'EPAP':
['totpop2010','tothhs2010','p_wrkage','pct_autoo0','totemp2010','totwrk2010',
'arealnd_ac','areaunp_ac','popdens_ac','empdens_ac','empentropy','trpequilib','ntwkdenped',
'intrsndens','empbytrans','jobacc45tr','retailempl','forst_nlcd','natrl_nlcd','treec_nlcd',
'opens_nlcd','pct_childr','pct_adults','pct_senior','pct_popwht','pct_popmal','pct_popfem','pct_nohsed',
'pct_hseduc','pct_2ycoll','pct_4ypcol','pct_worker','pct_notwrk','pct_hhkids','pct_nokids',
'pct_ownocc','pct_rentoc','avg_hhsize','pct_lowinc','pct_medinc','pct_higinc','topenspace']
}


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "UD4H HIAA Tools"
        self.alias = "UD4H HIAA Tools"

        # List of tool classes associated with this toolbox
        self.tools = [EPAPPost]


class EPAPPost(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = _hmv + " Data Request"
        self.description = "Provide HIAA API with " + _hmv + " data and receive health outcomes. " + \
                            "Supplying only a GEOID10 will return baseline health outcomes.  " + \
                            "Supplying other fields recognized by the " + _hmv + " schema will return custom health outcomes."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        # http://pro.arcgis.com/en/pro-app/arcpy/geoprocessing_and_python/defining-parameter-data-types-in-a-python-toolbox.htm
        # http://pro.arcgis.com/en/pro-app/arcpy/geoprocessing_and_python/accessing-parameters-within-a-python-toolbox.htm
        # http://pro.arcgis.com/en/pro-app/arcpy/geoprocessing_and_python/understanding-validation-in-script-tools.htm
        # http://desktop.arcgis.com/en/arcmap/10.3/analyze/creating-tools/customizing-script-tool-behavior.htm

        # this setting means if user selects existing file / feature class it will be overwritten
        arcpy.env.overwriteOutput = True

        # Parameter 0: The input ArcMap layer that contains the GEOID10s
        param_input = arcpy.Parameter(
            displayName="Input Census Block Group Features - must contain a CBG unique identifier field (GEOID10 or geoid10cbg)",
            name="in_features",
            datatype= "GPTableView",      # ["DETable", "GPFeatureLayer"],
            parameterType="Required",
            direction="Input")

        # Parameter 1: The feature class to write the API detail results to
        param_outfcdet = arcpy.Parameter(
            displayName="Output DETAIL Feature Class",
            name="outfcdet",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Output")

        # Parameter 2: The feature class to write the API summary results to
        param_outfcsum = arcpy.Parameter(
            displayName="Output SUMMARY Feature Class - leave blank to skip SUMMARY feature class creation",
            name="outfcsum",
            datatype="DEFeatureClass",
            parameterType="Optional",
            direction="Output")

        # Parameter 3: The base URL of the API HTTP request
        param_epaturl = arcpy.Parameter(
            displayName= _hmv + " Data Request URL",
            name="base_url",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        param_epaturl.value = _hiaa_url_default

        # Parameter 4: Email address associated with the client request
        param_client = arcpy.Parameter(
            displayName="Client Identifier - registered email address associated with this request",
            name="client_id",
            datatype="GPString",
            parameterType="Required",
            direction="Input")
        param_client.value = _clientid

        # Parameter 5: Assert only baseline values returned regardless of input schema
        param_baseonly = arcpy.Parameter(
            displayName="Return baseline data only (no custom inputs/outcomes, regardless of input schema)",
            name="baseonly",
            datatype="GPBoolean",
            parameterType="Required",
            direction="Input")
        param_baseonly.value = False

        params = [param_input,param_outfcdet,param_outfcsum,param_epaturl,param_client,param_baseonly]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        #return
        arcpy.env.overwriteOutput = True
        pass

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        #return
        arcpy.env.overwriteOutput = True

        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        """
            Steps:
                - Set out_jsonfile and dest_fc
                - Concatenate client_id to Base URL parameter
                - Output input_fc to json file for subsequent post
                - Make urllib2 request to API, save result in response
                - import response to geodb
        """


        ########################################################################
        # Validate the input fc for geoid10 and non-empty

        in_fc = parameters[0].valueAsText   # param_input

        record_count = int(arcpy.management.GetCount(in_fc)[0])
        arcpy.AddMessage("Selected row count in " + in_fc + " = " + str(record_count))
        if record_count == 0:
            messages.addErrorMessage("")
            messages.addErrorMessage("CANCELLING " + _hmv + " REQUEST: No input features selected")
            raise arcpy.ExecuteError

        epapfield_list = _epap_fields[_hmv]
        out_flist = []
        in_fc_fields = arcpy.ListFields(in_fc)
        geoidfn = ""
        for in_fc_field in in_fc_fields:
            if in_fc_field.name in ["GEOID10","geoid10cbg","geoid10c"]:
                out_flist.append(in_fc_field.name)
                geoidfn = in_fc_field.name
            elif in_fc_field.name in epapfield_list:
                out_flist.append(in_fc_field.name)

        if len(geoidfn) == 0:
            messages.addErrorMessage("")
            messages.addErrorMessage("CANCELLING " + _hmv + " REQUEST: input source has no field named 'GEOID10' or 'geoid10cbg'")
            raise arcpy.ExecuteError


        arcpy.AddMessage("Sending only these recognized fields to API: " + str(out_flist))

        ########################################################################
        # Output input_fc to json file for subsequent post
        # https://github.com/jasonbot/geojson-madness
        # http://stackoverflow.com/questions/19439961/python-requests-post-json-and-file-in-single-request
        # http://stackoverflow.com/questions/9746303/how-do-i-send-a-post-request-as-a-json
        import geojson_in
        import geojson_out

        dest_fc = parameters[1].valueAsText     # param_outfcdet

        outfc_path = os.path.dirname(dest_fc) #THY
        if outfc_path[-4:] == ".gdb":
            out_jsonpath = os.path.abspath(os.path.join(outfc_path,'..'))

        out_jsonfile = os.path.abspath(os.path.join(out_jsonpath, 'hiaa_prepost.json'))
        arcpy.AddMessage("out_jsonfile: " + "Export JSON file is " + str(out_jsonfile))

        geojson_out.write_geojson_table_file(in_fc, out_jsonfile, out_flist)

        arcpy.AddMessage(str(parameters[0].name) + ": Exported selected features of ArcMap layer " + str(in_fc) + " to custom json file '" + str(out_jsonfile) + "'")


        ########################################################################
        # Issue POST request using json from out_jsonfile

        client_id = parameters[4].valueAsText   # param_client
        base_only_bool = parameters[5].valueAsText     # param_baseonly
        base_only_int = '0'
        if base_only_bool == 'true':
            base_only_int = '1'
        in_geojson_url = parameters[3].valueAsText  # param_epaturl
        arcpy.AddMessage(_hmv + " Data Request URL: " + in_geojson_url)
        arcpy.AddMessage("URL Post Parameters: ClientID=" + client_id + ", BaseOnly=" + base_only_int)

        req = urllib2.Request(in_geojson_url)
        # req.add_header('Content-Type', 'application/json')

        with open(out_jsonfile, 'r') as myfile:
            file_json=myfile.read().replace('\n', '')

        post_json = {'postjson': file_json, 'clientid': client_id, 'baseonly': base_only_int}
        post_data = urllib.urlencode(post_json)
        arcpy.AddMessage("         >>>>>>>>>>>>>>>>>>>>>> COMMUNICATING WITH HIAA API  >>>>>>>>>>>>>>>>>>>>>>")
        try:
            response_handle = urllib2.urlopen(req, post_data )
        except:
            messages.addErrorMessage("")
            messages.addErrorMessage("API Communication Error.  It's possible the request was too large.  Try splitting it into multiple smaller requests.")
            raise arcpy.ExecuteError

        #arcpy.AddMessage(str(response_handle.info()))
        response_json = json.load(response_handle)
        # arcpy.AddMessage(json.dumps(response_json))

        # Check for error message returned
        try:
            return_err = str(response_json['Error'])
        except:
            return_err = ""

        if len(return_err) > 0:
            messages.addErrorMessage("")
            messages.addErrorMessage(_hmv + " DETAIL REQUEST FAILED.  Message from API: " + return_err)
            raise arcpy.ExecuteError

        # Check for no features due to bad geoid10s
        try:
            for item in response_json.get("features", []):
                break
        except:
            messages.addErrorMessage("")
            messages.addErrorMessage("Data Error: " + _hmv + " API could not find any census block group ids matching values in your input layer/table's '" + geoidfn + "' field")
            raise arcpy.ExecuteError


        # For potential SUMMARY below
        # need to get the requestid from the previously output json
        try:
            thisreqid = str(response_json['responseinfo']['requestid'])
            messages.AddMessage(_hmv + " detail request returned valid GeoJSON.  RequestID = " + thisreqid)
            req_proc_time = str(response_json['responseinfo']['request_processing_time'])
            messages.AddMessage(_hmv + " detail request processing time = " + req_proc_time)
            req_feat_count = str(response_json['responseinfo']['featurecount'])
            messages.AddMessage(_hmv + " API returned " + req_feat_count + " census block groups matched from total " + str(record_count) + " rows selected in input layer/table")
        except:
            thisreqid = "-1"

        ########################################################################
        # Set dest_fc and write to it
        arcpy.AddMessage(str(parameters[1].name) + ": " + "Output DETAIL feature class is " + str(dest_fc))

        out_schema = geojson_in.determine_schema(response_json)
        geojson_in.create_feature_class(dest_fc, out_schema)
        geojson_in.write_features(dest_fc, out_schema, response_json)
        arcpy.AddMessage("Saved API response to DETAIL output feature class " + dest_fc)


        ########################################################################
        # Make urllib2 request to API, save SUMMARY result in sum_handle
        dest_fcsum = parameters[2].valueAsText  # param_outfcsum
        if dest_fcsum is None:
            arcpy.AddMessage("Skipping creation of output SUMMARY feature class: none specified")
            create_fcsum = False
        else:
            arcpy.AddMessage(str(parameters[2].name) + ": " + "Output SUMMARY feature class is " + str(dest_fcsum))
            create_fcsum = True

        if create_fcsum and thisreqid <> "-1":

            base_sum_url =  in_geojson_url.replace(_hiaa_url_detail,_hiaa_url_summary)

            in_sum_url = base_sum_url + "?requestid=" + thisreqid
            in_sum_url = in_sum_url + "&clientid=" + client_id
            arcpy.AddMessage("Sending HTTP SUMMARY GET request: " + in_sum_url)

            arcpy.AddMessage("Importing SUMMARY JSON from URL")
            arcpy.AddMessage("         >>>>>>>>>>>>>>>>>>>>>> COMMUNICATING WITH API  >>>>>>>>>>>>>>>>>>>>>>")
            try:
                sum_handle = urllib2.urlopen(in_sum_url)
            except:
                messages.addErrorMessage("")
                messages.addErrorMessage("API Communication Error.  It's possible the request was too large.  Try splitting it into multiple smaller requests.")
                raise arcpy.ExecuteError

            response_json = json.load(sum_handle)


            # Check for error message returned
            try:
                return_err = str(response_json['Error'])
            except:
                return_err = ""

            if len(return_err) > 0:
                messages.addErrorMessage("")
                messages.addErrorMessage(_hmv + " SUMMARY REQUEST FAILED. Message from API: " + return_err)
                raise arcpy.ExecuteError

            # Check for no features due to bad geoid10s
            try:
                for item in response_json.get("features", []):
                    break
            except:
                messages.addErrorMessage("")
                messages.addErrorMessage("Data Error: " + _hmv + " API could not find a merge polygon matching your requestid.")
                raise arcpy.ExecuteError


            # For potential SUMMARY below
            # need to get the requestid from the previously output json
            try:
                thisreqid = str(response_json['responseinfo']['requestid'])
                messages.AddMessage(_hmv + " summary request returned valid GeoJSON.  RequestID = " + thisreqid)
                req_proc_time = str(response_json['responseinfo']['request_processing_time'])
                messages.AddMessage(_hmv + " summary request processing time = " + req_proc_time)
                req_feat_count = str(response_json['responseinfo']['featurecount'])
                messages.AddMessage(_hmv + " API returned summary feature with attributes.")
            except:
                thisreqid = "-1"

            out_schema = geojson_in.determine_schema(response_json)
            geojson_in.create_feature_class(dest_fcsum, out_schema)
            geojson_in.write_features(dest_fcsum, out_schema, response_json)
            arcpy.AddMessage("Saved URL response to SUMMARY output feature class " + dest_fcsum)


        return
