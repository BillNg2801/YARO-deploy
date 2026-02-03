const mongoose = require('mongoose');
const { graphFetch } = require('./graphClient');

const SUBSCRIPTIONS_COLLECTION = 'subscriptions';

function getNotificationUrl() {
  const base = process.env.BASE_URL || 'https://yarodeploy.vercel.app';
  return `${base.replace(/\/$/, '')}/api/webhook/mail`;
}

function getExpirationDateTime() {
  const d = new Date();
  d.setMinutes(d.getMinutes() + 4230);
  return d.toISOString();
}

async function createMailSubscription() {
  const notificationUrl = getNotificationUrl();
  const expirationDateTime = getExpirationDateTime();

  const body = {
    changeType: 'created',
    notificationUrl,
    resource: "me/mailFolders('Inbox')/messages",
    expirationDateTime,
    clientState: 'yaro-mail-subscription',
  };

  const result = await graphFetch('/subscriptions', {
    method: 'POST',
    body: JSON.stringify(body),
  });

  if (result?.id) {
    await saveSubscription(result);
  }
  return result;
}

async function saveSubscription(sub) {
  try {
    if (mongoose.connection.readyState !== 1) return;
    await mongoose.connection.db.collection(SUBSCRIPTIONS_COLLECTION).updateOne(
      { _id: 'mail_inbox' },
      {
        $set: {
          subscriptionId: sub.id,
          expirationDateTime: sub.expirationDateTime,
          resource: sub.resource,
          updatedAt: new Date(),
        },
      },
      { upsert: true }
    );
  } catch (err) {
    console.error('Failed to save subscription:', err.message);
  }
}

async function getStoredSubscription() {
  try {
    if (mongoose.connection.readyState !== 1) return null;
    return await mongoose.connection.db
      .collection(SUBSCRIPTIONS_COLLECTION)
      .findOne({ _id: 'mail_inbox' });
  } catch (err) {
    console.error('Failed to get subscription:', err.message);
    return null;
  }
}

async function renewSubscription(subscriptionId) {
  const expirationDateTime = getExpirationDateTime();
  await graphFetch(`/subscriptions/${subscriptionId}`, {
    method: 'PATCH',
    body: JSON.stringify({ expirationDateTime }),
  });
  const doc = await mongoose.connection.db
    .collection(SUBSCRIPTIONS_COLLECTION)
    .findOne({ _id: 'mail_inbox' });
  if (doc) {
    doc.expirationDateTime = expirationDateTime;
    doc.updatedAt = new Date();
    await mongoose.connection.db
      .collection(SUBSCRIPTIONS_COLLECTION)
      .updateOne({ _id: 'mail_inbox' }, { $set: doc });
  }
}

async function renewExpiringSubscriptions() {
  const doc = await getStoredSubscription();
  if (!doc?.subscriptionId) return;

  const exp = new Date(doc.expirationDateTime);
  const now = new Date();
  const hoursLeft = (exp - now) / (1000 * 60 * 60);
  if (hoursLeft < 24) {
    await renewSubscription(doc.subscriptionId);
    console.log('Renewed mail subscription');
  }
}

module.exports = {
  createMailSubscription,
  renewSubscription,
  renewExpiringSubscriptions,
  getStoredSubscription,
};
