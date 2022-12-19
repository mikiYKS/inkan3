  var authenticator;
  var client_id = "b52f7c0f-9962-4372-b84a-53bd6831df55";
  var redirect_url = "https://mikiyks.github.io/inkan3/";
  var scope;
  var access_token;

$(document).ready(function () {
  $("#run").click(() => tryCatch(getKakuin));
});

function getKakuin() {
  scope = "https://graph.microsoft.com/Files.Read.All";
  authenticator = new OfficeHelpers.Authenticator();
  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });
  authenticator
    .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
    .then(function (token) {
      access_token = token.access_token;
      //API呼び出し
      $(function () {
        $.ajax({
          url:
            "https://graph.microsoft.com/v1.0/drives/b!wh9TIKvGHk6lMpyOFa_tDYf465lhTsBJrcMIb7Y6agGXDdHPbUcwQagcaAlwEK_B/items/01SG44IHMJY6HM4OB2XJGZ34EYB77ZANB2",
          type: "GET",
          beforeSend: function (xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
          }
        }).then(
          async function (data) {
            const obj = data["@microsoft.graph.downloadUrl"];
            var kakuinbase64 = await getImageBase64(obj);

            //ここからkakuinbase64を張り付ける処理
            inkanpaste(kakuinbase64);
            //ログ出力
            $(function () {
              Excel.run(async (context) => {
                var inkanName = '角印';
                context.workbook.load("name");
                await context.sync();
                const alligator = ["XLSX", "XLSM", "XLSB", "XLS", "XLTX", "XLTM", "XLT"];
                const ext = context.workbook.name.split('.').pop().toUpperCase();
                if (alligator.indexOf(ext) == -1) {
                  var fileName = '未保存エクセル';
                } else {
                  var fileName = context.workbook.name;
                };
                inkanLog(inkanName, fileName);
              });
            });
          },
          function (data) {
            console.log(data);
          }
        );
      });
    })
    .catch(OfficeHelpers.Utilities.log);
}

// バイナリ画像をbase64で返す
async function getImageBase64(url) {
  const response = await fetch(url);
  const contentType = response.headers.get("content-type");
  const arrayBuffer = await response.arrayBuffer();
  let base64String = btoa(String.fromCharCode.apply(null, new Uint8Array(arrayBuffer)));
  //return `data:${contentType};base64,${base64String}`;
  return base64String;
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

Office.initialize = function (reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
};

//アクティブセルに印影貼り付け
async function inkanpaste(pic) {
  await Excel.run(async (context) => {
    const shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
    const cell = context.workbook.getActiveCell();
    cell.load("left").load("top");
    await context.sync();
    const shpStampImage = shapes.addImage(pic);
    shpStampImage.name = "印鑑";
    shpStampImage.scaleHeight(0.5, "OriginalSize");
    shpStampImage.left = cell.left;
    shpStampImage.top = cell.top;
    await context.sync();
  });
}

//SharePointListにログ出力
function inkanLog(inkanName, inkanFile) {
  scope = "https://graph.microsoft.com/Sites.ReadWrite.All";
  authenticator = new OfficeHelpers.Authenticator();
  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });
  //認証
  authenticator.authenticate(OfficeHelpers.DefaultEndpoints.Microsoft).then(function (token) {
    access_token = token.access_token;
    //API呼び出し印鑑ログ投稿
    $(function () {
      $.ajax({
        url:
          "https://graph.microsoft.com/v1.0/sites/20531fc2-c6ab-4e1e-a532-9c8e15afed0d/lists/6aac0560-622e-4ee1-ba8f-73b32d8e9f05/items",
        type: "POST",
        data: JSON.stringify({
          fields: {
            Title: inkanName,
            FileName: inkanFile
          }
        }),
        contentType: "application/json",
        beforeSend: function (xhr) {
          xhr.setRequestHeader("Authorization", "Bearer " + access_token);
        }
      }).then(
        async function (data) { },
        function (data) {
          console.log(data);
        }
      );
    });
  });
}
