/*
MIT License

Copyright(c) 2017 Chris Stayte

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace SyncArcProToGoogleEarth
{
    internal class AP2GE : Button
    {
        // Location of save file and file names
        static private String _saveDirectory;
        static private String _currentViewFileName = "AP2GE_CurrentView.kml";
        static private String _networkLinkFileName = "AP2GE_NetworkLink.kml";

        // Default Location Is Batman Building In Japan
        private String _latitude = "26.357896";
        private String _longitude = "127.783809";
        private String _altitude = "100";
        private String _heading = "0";
        private String _tilt = "0";

        #region Button 

        private AP2GE()
        {
            _saveDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"AP2GE\");
        }

        protected override void OnClick()
        {
            if (this.IsChecked)
            {
                MapViewCameraChangedEvent.Unsubscribe(MapViewCameraCanged);
                this.IsChecked = false;
                this.Caption = "Activate";

                string currentViewFile = Path.Combine(_saveDirectory, _currentViewFileName);

                if (File.Exists(currentViewFile))
                {
                    File.Delete(currentViewFile);
                }
            }
            else
            {
                MapViewCameraChangedEvent.Subscribe(MapViewCameraCanged, false);
                WriteCurrentView();
                WriteNetworkLink();
                this.IsChecked = true;
                this.Caption = "Deactivate";
            }
        }

        #endregion

        #region Methods

        private void MapViewCameraCanged(MapViewCameraChangedEventArgs args)
        {
            MapView mapView = args.MapView;

            if (mapView != null)
            {
                SyncViews(mapView);
            }
        }

        private async void SyncViews(MapView mapView)
        {
            // Get Altitude *Google Earth Range
            MapPoint lowerLeftPoint = null;
            MapPoint upperRightPoint = null;

            await QueuedTask.Run(() =>
            {
                lowerLeftPoint = MapPointBuilder.CreateMapPoint(mapView.Extent.XMin, mapView.Extent.YMin, mapView.Extent.SpatialReference);
                upperRightPoint = MapPointBuilder.CreateMapPoint(mapView.Extent.XMax, mapView.Extent.YMax, mapView.Extent.SpatialReference);
            });

            double latXmin, latXmax, longYmin, longYmax, diagonal;

            Tuple<double, double> tupleResult = null;

            tupleResult = await PointToLatLong(lowerLeftPoint);
            longYmin = tupleResult.Item1;
            latXmin = tupleResult.Item2;

            tupleResult = await PointToLatLong(upperRightPoint);
            longYmax = tupleResult.Item1;
            latXmax = tupleResult.Item2;

            diagonal = Math.Round(DistanceBetweenPoints(latXmin, longYmin, latXmax, longYmax, 'K') * 1000, 2); // 1km * 1000m

            double altitude = diagonal * Math.Sqrt(3) * 0.5;

            _altitude = altitude.ToString();



            // Longitude, Latitude
            MapPoint point = null;

            await QueuedTask.Run(() =>
            {
                Coordinate coordinate = new Coordinate(mapView.Camera.X, mapView.Camera.Y);
                //Coordinate coordinate = new Coordinate(((map.Extent.XMax + map.Extent.XMin) / 2), ((map.Extent.YMax + map.Extent.YMin) / 2));
                MapPointBuilder pointBuilder = new MapPointBuilder(coordinate, mapView.Extent.SpatialReference);
                point = pointBuilder.ToGeometry();
            });

            double latitude;
            double longitude;

            var result = await PointToLatLong(point);
            longitude = result.Item1;
            latitude = result.Item2;

            _latitude = Convert.ToString(Math.Round(latitude, 5));
            _longitude = Convert.ToString(Math.Round(longitude, 5));

            // Altitude
            // Arcmap Heading 0, 90, 180, -90:       Point of Reference is the Camera
            // Google Earth Heading 0, 90, 180, 270: Point of Reference is the Bearing
            double heading = mapView.Camera.Heading;
            heading = 360 - (heading < 0 ? (heading + 360) : heading);  // Flip Point of Reference - Flip to 360 
            _heading = heading.ToString();  


            // Tilt
            double tilt = (mapView.Camera.Pitch + 90); // Convert To Google Earth Formatting
            _tilt = tilt.ToString();
            

            WriteCurrentView();
        }

        private void WriteNetworkLink()
        {
            if (!System.IO.Directory.Exists(_saveDirectory))
            {
                System.IO.Directory.CreateDirectory(_saveDirectory);
            }

            string networklinkfile = Path.Combine(_saveDirectory, _networkLinkFileName);

            using (TextWriter tw = new StreamWriter(networklinkfile))
            {
                tw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                tw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
                tw.WriteLine("<Folder>");
                tw.WriteLine("<open>1</open>");
                tw.WriteLine("<name>ArcPro to Google Earth Sync</name>");
                tw.WriteLine("<NetworkLink>");
                tw.WriteLine("<name>" + "ArcGIS Pro Map" + "</name>");
                tw.WriteLine("<flyToView>1</flyToView>");
                tw.WriteLine("<Link>");
                tw.WriteLine("<href>" +_currentViewFileName + "</href>");
                tw.WriteLine("<refreshMode>onInterval</refreshMode>");
                tw.WriteLine("<refreshInterval>0.300000</refreshInterval>");
                tw.WriteLine("</Link>");
                tw.WriteLine("</NetworkLink>");
                tw.WriteLine("</Folder>");
                tw.WriteLine("</kml>");
            }

            System.Diagnostics.Process.Start(networklinkfile);
        }

        private void WriteCurrentView()
        {
            if (!System.IO.Directory.Exists(_saveDirectory))
            {
                System.IO.Directory.CreateDirectory(_saveDirectory);
            }

            string currentviewfile = Path.Combine(_saveDirectory, _currentViewFileName);

            using (TextWriter tw = new StreamWriter(currentviewfile))
            {

                tw.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                tw.WriteLine("<kml xmlns=\"http://www.opengis.net/kml/2.2\">");
                tw.WriteLine("<NetworkLinkControl>");
                tw.WriteLine("<LookAt>");
                tw.WriteLine("<longitude>" + _longitude + "</longitude>");
                tw.WriteLine("<latitude>" + _latitude + "</latitude>");
                tw.WriteLine("<altitudeMode>relativeToGround</altitudeMode>");
                // tw.WriteLine("<altitude> + " + _gpsAltitude + "</altitude>");
                tw.WriteLine("<heading>" + _heading + "</heading>");
                tw.WriteLine("<tilt>" + _tilt + "</tilt>");
                tw.WriteLine("<range>" + _altitude + "</range>");
                tw.WriteLine("</LookAt>");
                tw.WriteLine("</NetworkLinkControl>");
                tw.WriteLine("</kml>");
            }
        }

        private async Task<Tuple<double, double>> PointToLatLong(MapPoint point)
        {

            try
            {
                SpatialReference spatialReference = null;
                await QueuedTask.Run(() =>
                {
                    SpatialReferenceBuilder spatialReferenceBuilder = new SpatialReferenceBuilder(SpatialReferences.WGS84);

                    spatialReferenceBuilder.FalseX = -180;
                    spatialReferenceBuilder.FalseY = -90;
                    spatialReferenceBuilder.XYScale = 1000000;

                    spatialReference = spatialReferenceBuilder.ToSpatialReference();
                }); 

                Geometry geometry = (Geometry)point;

                MapPoint newPoint = null;

                await QueuedTask.Run(() =>
                {
                    newPoint = GeometryEngine.Project(geometry, spatialReference) as MapPoint;
                });

                return new Tuple<double, double>(newPoint.X, newPoint.Y);
                
            } catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;
        }

        private double DistanceBetweenPoints(double lat1, double lon1, double lat2, double lon2, char unit)
        {
            //'M' is statute miles
            //'K' is kilometers (default)
            //'N' is nautical miles  
            double theta = lon1 - lon2;
            double dist = Math.Sin(ConvertDegrees2Radians(lat1)) * Math.Sin(ConvertDegrees2Radians(lat2)) + Math.Cos(ConvertDegrees2Radians(lat1)) * Math.Cos(ConvertDegrees2Radians(lat2)) * Math.Cos(ConvertDegrees2Radians(theta));
            dist = Math.Acos(dist);
            dist = ConvertRadians2Degrees(dist);
            dist = dist * 60 * 1.1515;
            if (unit == 'K')
            {
                dist = dist * 1.609344;
            }
            else if (unit == 'N')
            {
                dist = dist * 0.8684;
            }
            return (dist);
        }

        private double ConvertDegrees2Radians(Double deg)
        {
            return (deg * Math.PI / 180.0);
        }

        private double ConvertRadians2Degrees(Double rad)
        {
            return (rad / Math.PI * 180.0);
        }

        #endregion
    }
}
