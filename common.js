
/**
 * 判断字符串 或 元素是否为空
 * @param str
 * @returns {boolean}
 */
function isEmpty(str) {
    return str === undefined || str === null || str === "" || str.trim().length === 0
}


/**
 * 格式化日期
 * @param format yyyyMMdd
 * @returns {*}
 */
Date.prototype.format = function(format) {
    let o = {
        'M+': this.getMonth() + 1,
        'd+': this.getDate(),
        'h+': this.getHours(),
        'm+': this.getMinutes(),
        's+': this.getSeconds(),
        'q+': Math.floor((this.getMonth() + 3) / 3),
        'S': this.getMilliseconds()
    }
    let dateRes;
    if (/(y+)/.test(format)) {
        dateRes = format.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    }
    for (let k in o) {
        if (new RegExp("(" + k + ")").test(format)) {
            dateRes = format.replace(RegExp.$1, RegExp.$1.length === 1 ? o[k] : ("00" + o[k]).substr(("" + o[k]).length));
        }
    }
    return dateRes;
}

/**
 * 返回传入日期的上一个月的今天
 * @param date
 * @param format
 * @returns {*}
 */
function lastMonth(date, format) {
    let month = date.getMonth();
    let the_last_month;
    if (month === 0) {
        the_last_month = date.setMonth(11);
    } else {
        the_last_month = date.setMonth(month - 1);
    }
    if (!format) {
        format = 'yyyy-MM-dd'
    }
    the_last_month =  new Date(the_last_month).format(format);
    return the_last_month;
}

/**
 * 得到传入日期的昨天
 * @param date
 * @returns {number}
 */
function yesterday(date, format) {
    if (!format) {
        format = 'yyyy-MM-dd'
    }
    return date.setDate(date.getDate() - 1).format(format);
}


