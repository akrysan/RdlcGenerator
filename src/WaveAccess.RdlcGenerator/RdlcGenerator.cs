namespace WaveAccess.RdlcGenerator {
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Linq;
    using Microsoft.Reporting.WebForms;
    using iTextSharp.text.pdf;

    public class RdlcGenerator {
        private LocalReport _localReport;
        private Dictionary<string, MethodInfo> _dataSetSources;
        private Dictionary<string, Stream> _subReports;
        private Assembly _assembly;

        public Document Generate(Assembly assembly, string reportPath, NameValueCollection parameters, string format, bool calculatePageCount = false) {
            _assembly = assembly;
            LoadReport(reportPath);
            FillData(parameters);
            return Generate(format, calculatePageCount);
        }

        public Document Generate(Assembly assembly, string reportPath, IDictionary<string, string> parameters, string format, bool calculatePageCount = false) {
            var collection = new NameValueCollection();
            foreach (var kvp in parameters) {
                collection.Add(kvp.Key, kvp.Value);
            }
            return Generate(assembly, reportPath, collection, format, calculatePageCount);
        }

        private Document Generate(string format, bool calculatePageCount) {
            string mimeType, encoding, extension;
            Warning[] warn;
            string[] streamids;
            int pageCount = 0;
            _localReport.SubreportProcessing += new SubreportProcessingEventHandler(localReport_SubreportProcessing);
            byte[] content = _localReport.Render(format, null, out mimeType, out encoding, out extension, out streamids, out warn);
            if (calculatePageCount && format == "pdf") {
                PdfReader pdfReader = new PdfReader(content);
                pageCount = pdfReader.NumberOfPages;
            }

            return new Document(content, mimeType, encoding, extension, streamids, pageCount);
        }

        private void FillSubReports(Stream stream, string reportNamespace) {
            XNamespace rd = "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner";
            stream.Position = 0;
            var rpt = XDocument.Load(stream);
            XNamespace rn = rpt.Root.GetDefaultNamespace();
            var subList = (from c in rpt.Descendants(rn + "ReportName") select c).Select(x => x.Value).ToList();
            _subReports = new Dictionary<string, Stream>();
            foreach (var item in subList.Distinct()) {
                _subReports.Add(item, _assembly.GetManifestResourceStream($"{reportNamespace}.{item}.rdlc"));
            }
        }

        private void LoadReport(string reportPath) {
            Stream stream = _assembly.GetManifestResourceStream(reportPath + ".rdlc");
            using (MemoryStream memoryStream = new MemoryStream()) {
                stream.CopyTo(memoryStream);
                _localReport = new LocalReport();
                memoryStream.Position = 0;
                _localReport.LoadReportDefinition(memoryStream);
                _dataSetSources = new Dictionary<string, MethodInfo>();

                var reportNamespace = reportPath.Substring(0, reportPath.LastIndexOf('.'));
                FillSubReports(stream, reportNamespace);

                if (_subReports != null) {
                    foreach (var subReport in _subReports) {
                        _localReport.LoadSubreportDefinition(subReport.Key, subReport.Value);
                        FillDataSources(subReport.Value);
                    }
                }

                _localReport.EnableHyperlinks = true;

                FillDataSources(memoryStream);
            }
        }

        private void FillDataSources(Stream stream) {
            XNamespace rd = "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner";
            stream.Position = 0;
            var rpt = XDocument.Load(stream);
            XNamespace rn = rpt.Root.GetDefaultNamespace();// "http://schemas.microsoft.com/sqlserver/reporting/2008/01/reportdefinition";
            if (rpt.Element(rn + "Report").Element(rn + "DataSets") == null) {
                return;
            }
            var dsets = from ds in rpt.Element(rn + "Report").Element(rn + "DataSets").Elements(rn + "DataSet")
                        select new {
                            Name = ds.Attribute("Name").Value,
                            Type = ds.Element(rd + "DataSetInfo").Element(rd + "ObjectDataSourceType").Value,
                            Sign = ds.Element(rd + "DataSetInfo").Element(rd + "ObjectDataSourceSelectMethodSignature").Value
                        };


            foreach (var dataset in dsets) {
                var tpS = Type.GetType(dataset.Type);

                var mi = tpS.GetMethods().FirstOrDefault(m => m.ToString() == dataset.Sign) as MethodInfo;
                _dataSetSources.Add(dataset.Name, mi);
            }
        }

        private void localReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e) {
            NameValueCollection parameters = new NameValueCollection();
            foreach (var p in e.Parameters) {
                parameters.Add(p.Name, p.Values[0]);
            }

            foreach (var dsName in e.DataSourceNames) {
                e.DataSources.Add(GetDataSource(dsName, parameters));
            }
        }

        private ReportDataSource GetDataSource(string dsName, NameValueCollection parameters) {
            var mi = _dataSetSources[dsName];
            var arrParams = mi.GetParameters();
            var paramsValue = new object[arrParams.Length];
            for (int i = 0; i < arrParams.Length; i++) {
                var pi = arrParams[i];
                var pType = pi.ParameterType;
                object value = null;
                var key = parameters.AllKeys.FirstOrDefault(p => p.ToLower() == pi.Name.ToLower());
                if (!string.IsNullOrEmpty(key)) {
                    var vals = parameters.GetValues(key);
                    value = Convert.ChangeType(vals[0], pType);
                }
                paramsValue[i] = value;
            }
            var ds = mi.Invoke(Activator.CreateInstance(mi.DeclaringType), paramsValue);
            return new ReportDataSource(dsName, ds);
        }

        private void FillData(NameValueCollection parameters) {
            foreach (var dsName in _localReport.GetDataSourceNames()) {
                _localReport.DataSources.Add(GetDataSource(dsName, parameters));
            }
            List<ReportParameter> prms = new List<ReportParameter>();
            foreach (var pi in _localReport.GetParameters()) {
                var key = parameters.AllKeys.FirstOrDefault(p => p.ToLower() == pi.Name.ToLower());
                string[] value = null;
                if (!string.IsNullOrEmpty(key)) {
                    value = parameters.GetValues(key) ?? new string[] { "" };
                    prms.Add(new ReportParameter(pi.Name, value));
                }
            }
            _localReport.SetParameters(prms);
        }
    }
}
