namespace QHM
{
    public class Data
    {
        public List<MyPoint>? MyPoints { get; set; }
        public List<int[]> Tb { set;get; } = new List<int[]>();
    }
    public class  MyPoint
    {
        public int x { set; get; }
        public int y { set; get; }
        public int index { set; get; }
        public int W { set; get; }
        public bool IsAccess { set; get; }
        public bool IsBackbone { set; get; }
        public MyPoint? backbone { set; get; }
        public double? moment { set; get; }
        public double? backboneDistance { set; get; }
        
    }
   
   
}
