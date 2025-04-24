SPFarm objFarm = SPFarm.Local;
            long farmVersion = objFarm.Version;
            Guid farmID = objFarm.Id;
            SPObjectStatus farmStatus = objFarm.Status;
            Console.WriteLine(“Farm Version is:” + farmVersion);
            Console.WriteLine(“Farm ID is:” + farmID);
            Console.WriteLine(“Farm Status is:” + farmStatus);
            Console.ReadKey();