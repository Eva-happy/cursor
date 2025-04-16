// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const { serviceId, amount } = event
  const wxContext = cloud.getWXContext()
  const userId = wxContext.OPENID
  
  try {
    // 创建订单记录
    const orderResult = await db.collection('orders').add({
      data: {
        userId,
        serviceId,
        amount,
        status: 'pending', // pending, paid, cancelled
        createdAt: db.serverDate(),
        updatedAt: db.serverDate()
      }
    })
    
    // 生成订单号
    const orderId = orderResult._id
    const timeStamp = new Date().getTime()
    const nonceStr = Math.random().toString(36).substr(2, 15)
    
    // 调用支付接口获取支付参数
    const res = await cloud.cloudPay.unifiedOrder({
      body: '玄学服务咨询费用',
      outTradeNo: orderId,
      spbillCreateIp: '127.0.0.1',
      subMchId: '1900000109', // 替换为您的微信支付商户号
      totalFee: amount * 100, // 金额转为分
      envId: cloud.DYNAMIC_CURRENT_ENV,
      functionName: 'payCallback'
    })
    
    // 返回支付参数
    return {
      payment: {
        timeStamp,
        nonceStr,
        package: `prepay_id=${res.payment.prepayId}`,
        signType: 'MD5',
        paySign: res.payment.paySign,
        orderId
      }
    }
    
  } catch (err) {
    console.error(err)
    return {
      error: err.message
    }
  }
} 