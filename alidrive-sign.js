// 当前日期
var currentDate = new Date();
// 年
var currentYear = currentDate.getFullYear();
// 月
var currentMonth = currentDate.getMonth() + 1;
// 日
var currentDay = currentDate.getDate();
// 当月最后一天
var lastDayOfMonth = new Date(currentYear, currentMonth, 0).getDate();
// 当前日期字符串
var dateString = currentYear + '年' + currentMonth + '月' + currentDay + "日";

// 输出日志
function log(message) {
    console.log(message);
}

// 发送通知
function notify(title, messages, notifyType, notifyParams) {
    log(notifyType)
    const notificationMethods = {
        'bark': barkNotify,
        'pushplus': pushplusNotify,
        // 添加其他通知方式的映射
    };

    for (const method in notificationMethods) {
        if (notifyType.includes(method)) {
            notificationMethods[method](title, messages, notifyParams);
            break;
        }
    }
}

// pushplus通知
function pushplusNotify(title, messages, token) {
    messages = messages.replace(/\//g, "-");
    HTTP.post("https://www.pushplus.plus/send/", {
        "token": token,
        "title": title,
        "content": messages,
    })
}

// Bark 通知
function barkNotify(title, messages, url) {
    url = url.replace(/\/这里改成你自己的推送内容/g, "")
    messages = messages.replace(/\//g, "-");
    title = encodeURIComponent(title)
    messages = encodeURIComponent(messages)
    let urls = messages === "" ? `${url}/${title}` : `${url}/${title}/${messages}`;
    HTTP.get(urls)
}

// token列
var tokenColumn = "A"
// 是否跳过列
var skipColumn = "B"
// 通知类型列
var notifyTypeColumn = "C"
// 通知参数列
var notufyParamsColumn = "D"
// 当月签到list原始数据
var logColumn = "E"
// 每日签到奖励原始数据列
var signInRewardColumn = "F"
// 每日限时任务原始数据列
var signInRewardTaskColumn = "G"



// 遍历行
for (let row = 2; row <= 20; row++) {
    // 获取刷新令牌
    var refresh_token = Application.Range(tokenColumn + row).Text;
    // 获取跳过列
    var skip = Application.Range(skipColumn + row).Text;
    // 获取通知类型
    var notifyType = Application.Range(notifyTypeColumn + row).Text;
    // 获取通知参数
    var notifyParams = Application.Range(notufyParamsColumn + row).Text

    if (skip === "是") {
        log("单元格【" + skipColumn + row + "】内的值为是，跳过本次签到");
        continue;
    } else {
        // 判断刷新令牌是否为空
        if (refresh_token != "") {
            // 通过刷新令牌获取访问令牌
            let data = HTTP.post("https://auth.aliyundrive.com/v2/account/token",
                JSON.stringify({
                    "grant_type": "refresh_token",
                    "refresh_token": refresh_token
                })
            );
            data = data.json();
            var access_token = data['access_token'];
            var phone = data["user_name"];
            var newRefreshToken = data["refresh_token"]
            // 更新刷新令牌
            Application.Range(tokenColumn + row).Value = newRefreshToken

            // 判断访问令牌是否获取成功
            if (access_token == undefined) {
                // 签到失败，输出错误信息并发送通知
                log("单元格【" + tokenColumn + row + "】内的token值错误，程序执行失败，请重新复制正确的token值");
                notify("阿里云盘签到失败", "单元格【" + tokenColumn + row + "】内的token值错误，程序执行失败，请重新复制正确的token值", notifyType, notifyParams)
                continue;
            }

            // 构造访问令牌
            var access_token2 = 'Bearer ' + access_token;
            // 获取签到信息
            let data2 = HTTP.post("https://member.aliyundrive.com/v2/activity/sign_in_list?_rx-s=mobile",
                JSON.stringify({}), {
                    headers: {
                        "Authorization": access_token2,
                    }
                }
            );
            data2 = data2.json(); // 将响应数据解析为 JSON 格式
            // 存储签到结果
            Application.Range(logColumn + row).Value = JSON.stringify(data2)
            var is_sign_in = data2['result']['isSignIn'];
            var signIn_count = data2['result']['signInCount']; // 获取签到次数
            var signInInfos = data2["result"]["signInInfos"]
            // 遍历签到信息
            for (var i = 0; i < signInInfos.length; i++) {
                // 判断签到日期是否与当前日期相同
                if (signIn_count.toString() === signInInfos[i]["day"]) {
                    // 遍历签到奖励
                    for (var j = 0; j < signInInfos[i]["rewards"].length; j++) {
                        // 判断签到奖励类型
                        if (signInInfos[i]["rewards"][j]["type"] === "dailySignIn") {
                            var todaySign = signInInfos[i]["rewards"][j]
                            var todaySignReward = todaySign["name"]
                            var todaySignRequire = todaySign["remind"]
                            var todaySignInStatus = todaySign["status"]
                            console.log("今日签到任务奖励：" + todaySignReward)
                            console.log("今日签到任务要求：" + todaySignRequire)
                        }
                        // 判断限时任务奖励
                        if (signInInfos[i]["rewards"][j]["type"] === "dailyTask") {
                            var todayTask = signInInfos[i]["rewards"][j]
                            var todayTaskReward = todayTask["name"]
                            var todayTaskRequire = todayTask["remind"]
                            var todayTaskStatus = todayTask["status"]
                            console.log("今日限时任务奖励：" + todayTaskReward)
                            console.log("今日限时任务要求：" + todayTaskRequire)
                        }
                    }
                }
            }
            // 构造签到日志和奖励信息
            var logMessage = `本月：签到${signIn_count}天\n日期：${dateString}`;
            var rewardMessage = `账号：${phone}\n`;
            // 判断签到状态
            if (todaySignInStatus === "verification") {
                rewardMessage += `签到：${todaySignReward} (已经领取)\n`;
            } else {
                try {
                    // 领取签到奖励
                    let data3 = HTTP.post(
                        "https://member.aliyundrive.com/v1/activity/sign_in_reward?_rx-s=mobile",
                        JSON.stringify({
                            "signInDay": signIn_count
                        }), {
                            headers: {
                                "Authorization": access_token2
                            }
                        }
                    );
                    data3 = data3.json();
                    Application.Range(signInRewardColumn + row).Value = JSON.stringify(data3)
                    // 判断签到奖励是否领取成功
                    if (data3["result"]) {
                        rewardMessage += `签到：${todayTaskReward} (领取完毕)\n`;
                    }
                } catch (error) {
                    notify("阿里云盘签到失败", error + "\n" + JSON.stringify(data3), notifyType, notifyParams)
                }
            }
            // 判断限时任务状态
            if (todayTaskStatus === "verification") {
                rewardMessage += `限时：${todayTaskReward} (已经领取)\n`;
            } else if (todayTaskStatus === "finished") {
                // 确定满足, 开始领取
                try {
                    let data4 = HTTP.post(
                        "https://member.aliyundrive.com/v2/activity/sign_in_task_reward?_rx-s=mobile",
                        JSON.stringify({
                            "signInDay": signIn_count
                        }), {
                            headers: {
                                "Authorization": access_token2
                            }
                        }
                    );
                    data4 = data4.json();
                    log(JSON.stringify(data4))
                    Application.Range(signInRewardTaskColumn + row).Value = JSON.stringify(data4)
                    // 判断领取是否成功
                    if (data4["result"]) {
                        rewardMessage += `限时：${todayTaskReward} (领取完毕)\n`;
                    }
                    // else {
                    //     rewardMessage += `限时：${todayTaskReward} (未知情况)\n要求：${todayTaskRequire}\nlogs：${JSON.stringify(data4)}\n`;
                    // }
                } catch (error) {
                    notify("阿里云盘任务失败", error + "\n" + JSON.stringify(data4), notifyType, notifyParams)
                    continue;
                }
            } else if (todayTaskStatus === "unfinished") {
                // 任务未完成
                rewardMessage += `限时：${todayTaskReward} (请做任务)\n要求：${todayTaskRequire}\n`;
            } else {
                // 未知情况
                rewardMessage += `限时：${todayTaskReward} (${todayTaskStatus})\n要求：${todayTaskRequire}\nlogs：${JSON.stringify(todayTask)}\n`;
            }
            // 输出签到日志和奖励信息
            log(rewardMessage + logMessage);
            notify("阿里云盘签到", rewardMessage + logMessage, notifyType, notifyParams)
        } else {
            1
        }
    }
}
