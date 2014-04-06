//var AcrowireViewModel = function (items) {
//    var self = this;
//    self.items = ko.observableArray(items);
//    self.repository = [];

//    self.PopulateArray = function () {
//        $.ajax({
//            url: "/Home/GetAcrowireReport?login=devpro01@acrowire.com&password=@cr0wire$",
//            async: true
//        }).done(function (data) {
//            self.items(data.GridList);
//            self.repository = data.GridList;
//        });
//    };

//    self.PopulateArray();
//};