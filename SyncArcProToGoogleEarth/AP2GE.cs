using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using ArcGIS.Desktop.Mapping.Events;
using System;
using System.IO;
using System.Threading.Tasks;

namespace SyncArcProToGoogleEarth
{
    internal class AP2GE : Button
    {
        // Location of save file and file names
        static private String _saveDirectory;
        static private String _currentViewFileName = "AP2GE_CurrentView.kml";
        static private String _networkLinkFileName = "AP2GE_NetworkLink.kml";

        private String _latitude = "";
        private String _longitude = "";
        private String _altitude = "";
        private String _heading = "0";
        private String _tilt = "0";

        private bool _cameraExisted = true;

        private enum enumLongLat
        {
            Latitude = 1,
            Longitude = 2
        };

        private enum enumReturnFormat
        {
            WithSigns = 0,
            NMEA = 1
        }

        #region Button 

        private AP2GE()
        {
            _saveDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"\AP2GE\");
        }

        protected override void OnClick()
        {
            if (this.IsChecked)
            {
                MapViewCameraChangedEvent.Unsubscribe(test);

                this.IsChecked = false;
            }
            else
            {
                MapViewCameraChangedEvent.Subscribe(test, false);

                WriteCurrentView();
                WriteNetworkLink();

                this.IsChecked = true;
            }
        }

        #endregion

        #region Methods

        private void test(MapViewCameraChangedEventArgs args)
        {
            MapView map = args.MapView;
            if (map != null)
            {
                SyncViews(map);
            }
        }

        private void drawStarted(MapViewEventArgs args)
        {
            MapView map = args.MapView;
            if (map == null) return;
            if (map.Camera != null)
            {
                SyncViews(map);
                return;
            }
            _cameraExisted = false;
        }

        private void drawCompleted(MapViewEventArgs args)
        {
            MapView map = args.MapView;
            if (map == null) return;
            if (!_cameraExisted)
            {
                SyncViews(map);
                _cameraExisted = true;
            }
           
        }

        private async void SyncViews(MapView map)
        {

            MapPoint lowerLeftPoint = null;
            MapPoint upperRightPoint = null;

            await QueuedTask.Run(() =>
            {
                lowerLeftPoint = MapPointBuilder.CreateMapPoint(map.Extent.XMin, map.Extent.YMin, map.Extent.SpatialReference);
                upperRightPoint = MapPointBuilder.CreateMapPoint(map.Extent.XMax, map.Extent.YMax, map.Extent.SpatialReference);
            });

            double latXmin;
            double latXmax;

            double longYmin;
            double longYmax;

            double diagonal;

            Tuple<double, double> tupleResult = null;

            tupleResult = await PointToLatLong(lowerLeftPoint);
            longYmin = tupleResult.Item1;
            latXmin = tupleResult.Item2;

            tupleResult = await PointToLatLong(upperRightPoint);
            longYmax = tupleResult.Item1;
            latXmax = tupleResult.Item2;

            diagonal = Math.Round(distance(latXmin, longYmin, latXmax, longYmax, 'K') * 1000, 2); // 1km * 1000

            _altitude = Convert.ToString(0.5 * Math.Sqrt(3) * diagonal);


            MapPoint point = null;

            await QueuedTask.Run(() =>
            {
                Coordinate coordinate = new Coordinate(((map.Extent.XMax + map.Extent.XMin) / 2), ((map.Extent.YMax + map.Extent.YMin) / 2));
                MapPointBuilder pointBuilder = new MapPointBuilder(coordinate, map.Extent.SpatialReference);
                point = pointBuilder.ToGeometry();
            });

            double latitude;
            double longitude;

            var result = await PointToLatLong(point);
            longitude = result.Item1;
            latitude = result.Item2;

            _latitude = Convert.ToString(Math.Round(latitude, 5));
            _longitude = Convert.ToString(Math.Round(longitude, 5));


            double heading = map.Camera.Heading;
            heading = heading < 0 ? (heading + 360) : heading;
            _heading = Convert.ToString(360 - heading);


            _tilt = map.Camera.Roll.ToString();

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

        private double distance(double lat1, double lon1, double lat2, double lon2, char unit)
        {
            //'M' is statute miles
            //'K' is kilometers (default)
            //'N' is nautical miles  
            double theta = lon1 - lon2;
            double dist = Math.Sin(degrees2radians(lat1)) * Math.Sin(degrees2radians(lat2)) + Math.Cos(degrees2radians(lat1)) * Math.Cos(degrees2radians(lat2)) * Math.Cos(degrees2radians(theta));
            dist = Math.Acos(dist);
            dist = radians2degrees(dist);
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

        private double degrees2radians(Double deg)
        {
            return (deg * Math.PI / 180.0);
        }

        private double radians2degrees(Double rad)
        {
            return (rad / Math.PI * 180.0);
        }

        #endregion
    }
}
