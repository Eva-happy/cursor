// 云函数入口文件
const cloud = require('wx-server-sdk')

cloud.init({
  env: cloud.DYNAMIC_CURRENT_ENV
})

const db = cloud.database()

// 云函数入口函数
exports.main = async (event, context) => {
  const { message } = event
  const wxContext = cloud.getWXContext()
  const userId = wxContext.OPENID
  
  try {
    // 记录用户消息
    await db.collection('messages').add({
      data: {
        userId,
        content: message,
        type: 'user',
        createdAt: db.serverDate()
      }
    })
    
    // 根据关键词匹配回复
    const reply = await generateReply(message)
    
    // 记录系统回复
    await db.collection('messages').add({
      data: {
        userId,
        content: reply,
        type: 'system',
        createdAt: db.serverDate()
      }
    })
    
    return {
      reply
    }
    
  } catch (err) {
    console.error(err)
    return {
      error: err.message
    }
  }
}

// 生成回复内容
async function generateReply(message) {
  // 关键词匹配规则
  const rules = [
    {
      keywords: ['价格', '费用', '多少钱'],
      reply: '我们提供多种服务套餐：\n1. 每日行程指南：9.9元\n2. 事业运势分析：19.9元\n3. 感情运势指导：19.9元\n4. 综合运势预测：39.9元\n具体服务内容可查看对应服务详情页面。'
    },
    {
      keywords: ['优惠', '折扣', '促销'],
      reply: '目前新用户首次购买任意服务可享受8折优惠。另外，购买年度套餐可享受7折优惠。详情请咨询客服。'
    },
    {
      keywords: ['退款', '退订', '取消'],
      reply: '如需退款，请在服务开始前提出申请。服务开始后将无法退款。具体退款规则请查看用户协议或联系客服。'
    },
    {
      keywords: ['生辰八字', '八字', '生辰'],
      reply: '在进行运势分析时，我们需要您的生辰八字信息（年月日时）。您可以在个人资料中填写，或在咨询开始时提供。'
    },
    {
      keywords: ['准确', '准确率', '可信'],
      reply: '我们的分析建议仅供参考，不作为决策依据。运势预测基于传统玄学理论，结合现代分析方法，为您提供参考建议。'
    }
  ]
  
  // 默认回复
  let reply = '感谢您的咨询。我是智能客服小助手，目前可以回答关于服务价格、优惠活动、退款规则等常见问题。如需更专业的解答，建议购买相关服务或联系人工客服。'
  
  // 遍历规则匹配关键词
  for (const rule of rules) {
    if (rule.keywords.some(keyword => message.includes(keyword))) {
      reply = rule.reply
      break
    }
  }
  
  return reply
} 