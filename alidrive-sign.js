var myDate = new Date(); // 创建一个表示当前时间的 Date 对象
var year = myDate.getFullYear();
var month = (myDate.getMonth() + 1).toString().padStart(2, '0'); // 月份是从0开始的，所以要加1
var day = myDate.getDate().toString().padStart(2, '0');

var data_time = year + '年' + month + '月' + day + "日"; // 获取当前日期的字符串表示

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d;); // 使程序暂停执行一段时间
}

function log(message) {
  console.log(message); // 打印消息到控制台
  // TODO: 将日志写入文件
}

function sendBark(title, messages, url) {
  //将消息中的/替换为-
  messages = messages.replace(/\//g, "-");
  title = encodeURIComponent(title)
  messages = encodeURIComponent(messages)
  log("Bark通知：title：" + title + "，messages：" + messages);
  if (messages === "") {
    let urls = url + "/" + title
    HTTP.get(urls)
  } else {
    let urls = url + "/" + title + "/" + messages
    HTTP.get(urls)
  }
}



// sendBark("阿里云盘签到开始", "")

var tokenColumn = "A"; // 设置列号变量为 "A"
var signInColumn = "B"; // 设置列号变量为 "B"
var rewardColumn = "C"; // 设置列号变量为 "C"
var barkColumn = "E"  //bark通知url列


for (let row = 2; row <= 20; row++) { // 循环遍历从第 2 行到第 20 行的数据

  var refresh_token = Application.Range(tokenColumn + row).Text; // 获取指定单元格的值
  var sflq = Application.Range(signInColumn + row).Text; // 获取指定单元格的值
  var sflqReward = Application.Range(rewardColumn + row).Text; // 获取指定单元格的值
  var barkUrl = Application.Range(barkColumn + row).Text


  if (sflq == "是") { // 如果“是否签到”为“是”
    if (refresh_token != "") { // 如果刷新令牌不为空
      // 发起网络请求-获取token
      let data = HTTP.post("https://auth.aliyundrive.com/v2/account/token",
        JSON.stringify({
          "grant_type": "refresh_token",
          "refresh_token": refresh_token
        })
      );
      data = data.json(); // 将响应数据解析为 JSON 格式
      var access_token = data['access_token']; // 获取访问令牌
      var phone = data["user_name"]; // 获取用户名
      var newRefreshToken = data["refresh_token"] //获取新refreshToken
      console.log(newRefreshToken)
      Application.Range(tokenColumn + row).Value = newRefreshToken

      if (access_token == undefined) { // 如果访问令牌未定义
        log("单元格【" + tokenColumn + row + "】内的token值错误，程序执行失败，请重新复制正确的token值");
        ///
        sendBark("阿里云盘签到失败", "单元格【" + tokenColumn + row + "】内的token值错误，程序执行失败，请重新复制正确的token值", barkUrl)
        ///
        continue; // 跳过当前行的后续操作
      }

      try {
        var access_token2 = 'Bearer ' + access_token; // 构建包含访问令牌的请求头
        // 签到
        let data2 = HTTP.post("https://member.aliyundrive.com/v1/activity/sign_in_list",
          JSON.stringify({
            "_rx-s": "mobile"
          }), {
          headers: {
            "Authorization": access_token2
          }
        }
        );
        data2 = data2.json(); // 将响应数据解析为 JSON 格式
        var signin_count = data2['result']['signInCount']; // 获取签到次数

        var logMessage = `本月：签到${signin_count}天\n账号：${phone}\n日期：${data_time}`;
        var rewardMessage = "";

        if (sflqReward == "是") { // 如果“是否领取奖励”为“是”
          if (sflq == "是") { // 如果“是否签到”为“是”
            try {
              // 领取奖励
              let data3 = HTTP.post(
                "https://member.aliyundrive.com/v1/activity/sign_in_reward?_rx-s=mobile",
                JSON.stringify({
                  "signInDay": signin_count
                }), {
                headers: {
                  "Authorization": access_token2
                }
              }
              );
              data3 = data3.json(); // 将响应数据解析为 JSON 格式
              var rewardName = data3["result"]["name"]; // 获取奖励名称
              var rewardDescription = data3["result"]["notice"]; // 获取奖励描述
              rewardMessage = `获得：${rewardName}(${rewardDescription})`;
            } catch (error) {
              if (error.response && error.response.data && error.response.data.error) {
                var errorMessage = error.response.data.error; // 获取错误信息
                if (errorMessage.includes(" - 今天奖励已领取")) {
                  rewardMessage = "状态：已领取";
                  log("账号：" + phone + " - " + rewardMessage);
                } else {
                  log("账号：" + phone + " - 奖励领取失败：" + errorMessage);
                }
              } else {
                log("账号：" + phone + " - 奖励领取失败");
              }
            }
          } else {
            rewardMessage = "状态：待领取";
          }
        } else {
          rewardMessage = "状态：待领取";
        }

        log(logMessage + rewardMessage);
        sendBark("阿里云盘签到", rewardMessage + "\n" + logMessage, barkUrl)
      } catch {
        log("单元格【" + tokenColumn + row + "】内的token签到失败");
        sendBark("阿里云盘签到", "单元格【" + tokenColumn + row + "】内的token签到失败", barkUrl)
        continue; // 跳过当前行的后续操作
      }
    } else {
      log("账号：" + phone + " 不签到");
    }
  }
}

var currentDate = new Date(); // 创建一个表示当前时间的 Date 对象
var currentDay = currentDate.getDate(); // 获取当前日期的天数
var lastDayOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0).getDate(); // 获取当月的最后一天的日期

if (currentDay === lastDayOfMonth) { // 如果当前日期是当月的最后一天
  log("强制领取本月奖励")
  for (let row = 2; row <= 20; row++) { // 循环遍历从第 2 行到第 20 行的数据
    var sflq = Application.Range(signInColumn + row).Text; // 获取指定单元格的值
    var sflqReward = Application.Range(rewardColumn + row).Text; // 获取指定单元格的值
    var barkUrl = Application.Range(barkColumn + row).Text;

    if (sflq === "是") { // 如果“是否签到”为“是”，则强制领取
      var refresh_token = Application.Range(tokenColumn + row).Text; // 获取指定单元格的值
      var jsyx = Application.Range(emailColumn + row).Text; // 获取指定单元格的值

      if (refresh_token !== "") { // 如果刷新令牌不为空
        // 发起网络请求-获取token
        let data = HTTP.post("https://auth.aliyundrive.com/v2/account/token",
          JSON.stringify({
            "grant_type": "refresh_token",
            "refresh_token": refresh_token
          })
        );
        data = data.json(); // 将响应数据解析为 JSON 格式
        var access_token = data['access_token']; // 获取访问令牌
        var phone = "账号：" + data["user_name"]; // 获取用户名
        if (access_token === undefined) { // 如果访问令牌未定义
          log("单元格【" + tokenColumn + row + "】内的token值错误，程序执行失败，请重新复制正确的token值");
          sendBark("阿里云盘签到", "领取全部奖励时，单元格【" + tokenColumn + row + "】内的token值错误，程序执行失败，请重新复制正确的token值", barkUrl)
          continue; // 跳过当前行的后续操作
        }
        var status = [];
        for (day = 1; day <= lastDayOfMonth; day++) {
          try {
            var access_token2 = 'Bearer ' + access_token; // 构建包含访问令牌的请求头
            // 领取奖励
            let data4 = HTTP.post(
              "https://member.aliyundrive.com/v1/activity/sign_in_reward?_rx-s=mobile",
              JSON.stringify({
                "signInDay": day
              }), {
              headers: {
                "Authorization": access_token2
              }
            }
            );
            data4 = data4.json(); // 将响应数据解析为 JSON 格式
            var claimStatus = data4["success"]; // 获取奖励状态
            if (claimStatus === false) {
              log("账号：" + phone + " - 第 " + day + " 天奖励领取失败");
              status.push(day)
            }
          } catch {
            log("单元格【" + tokenColumn + row + "】内的token签到失败");
            sendBark("阿里云盘签到", "领取全部奖励时，单元格【" + tokenColumn + row + "】内的token签到失败", barkUrl)
            continue; // 跳过当前行的后续操作
          }
        }
        if (status.length == 0) {
          log(phone + " - 本月奖励领取成功");
          sendBark("阿里云盘签到", phone + " 本月奖励领取成功", barkUrl)
        } else {
          var text = "";
          for (i = 0; i < status.length - 1; i++) {
            text += status[i] + "、";
          }
          text += status[status.length];
          log(phone + " - 除第 " + text + " 天奖励领取失败外，本月其余天数均成功");
          sendBark("阿里云盘签到", phone + " - 除第 " + text + " 天奖励领取失败外，本月其余天数均成功", barkUrl)
        }
      } else {
        log(phone + " 不签到");
      }
    }
  }
  log("自动领取未领取奖励完成。");
}
