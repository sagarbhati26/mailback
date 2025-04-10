import { Schema, model } from 'mongoose';

const emailSchema = new Schema({
    to: {
      type: String,
      required: true,
    },
    subject: {
      type: String,
      required: true,
    },
    message: {
      type: String,
      required: true,
    },
    sentAt: {
        type: Date,
        default: Date.now,
      }
  });
export default model('Email', emailSchema);